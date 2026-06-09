"""StoreKit 1 wrapper for In-App Purchases (donations + XML Pro subscription).

Why StoreKit 1 (not 2): StoreKit 2 (`Product`, `Transaction`, async/await) is
Swift-only and has no PyObjC binding. StoreKit 1 (`SKProductsRequest`,
`SKPaymentQueue`, `SKPaymentTransactionObserver`) is fully bound by
``pyobjc-framework-StoreKit`` and can purchase Auto-Renewable Subscriptions too.
See outputs/tabledown-monetization-plan.md §3.

Product layout (App Store Connect — IDs are immutable once registered):
- 3 donation tiers  -> Consumable, **no reward** (Apple requires consumables to
  grant nothing of lasting value).
- XML Pro           -> Auto-Renewable Subscription ($4.99/year), unlocks the XML
  conversion menu item + the ⌘⌃C global hotkey.

The module degrades gracefully: if ``StoreKit`` cannot be imported (e.g. a source
run where the framework isn't bundled) the app still starts and treats Pro as
active for development convenience — see ``Store.pro_active``. Nothing here ever
raises on the import path, so ``tablemark.app`` imports cleanly headless.
"""
from __future__ import annotations

import threading
import time

from Foundation import NSNumberFormatter, NSObject

from .logger import log
from . import settings

# --- Product IDs (exact, from the spec — immutable once registered in ASC) ---
TIP_SMALL = "com.tabledown.app.tip.small"
TIP_MEDIUM = "com.tabledown.app.tip.medium"
TIP_LARGE = "com.tabledown.app.tip.large"
PRO_YEARLY = "com.tabledown.app.pro.yearly"

TIP_PRODUCT_IDS = {
    "small": TIP_SMALL,
    "medium": TIP_MEDIUM,
    "large": TIP_LARGE,
}
ALL_PRODUCT_IDS = (TIP_SMALL, TIP_MEDIUM, TIP_LARGE, PRO_YEARLY)

# Shown in the subscription sheet before product metadata loads (or when the
# fetch failed). Must match the base price configured in App Store Connect;
# once metadata arrives the localized store price replaces it.
PRO_PRICE_FALLBACK = "$4.99"

# A subscription year, plus a small grace window so a brief offline spell or a
# slightly-late auto-renew transaction doesn't lock a paying user out mid-use
# (the util is intentionally lenient — see spec §3.5 / §9).
_YEAR_SECONDS = 365 * 24 * 60 * 60
_GRACE_SECONDS = 3 * 24 * 60 * 60

# StoreKit 1 transaction states (SKPaymentTransactionState*).
_STATE_PURCHASING = 0
_STATE_PURCHASED = 1
_STATE_FAILED = 2
_STATE_RESTORED = 3
_STATE_DEFERRED = 4

# SKError: the one failure that must stay silent in the UI (user backed out).
_SK_ERROR_DOMAIN = "SKErrorDomain"
_SK_ERROR_PAYMENT_CANCELLED = 2

# Product metadata fetch retry — the app fetches once at startup, which can
# race a not-yet-connected network. A couple of spaced retries keeps menu
# prices from staying blank for the whole session.
_MAX_FETCH_RETRIES = 2
_FETCH_RETRY_SECONDS = 20

# NSNumberFormatterCurrencyStyle — format SKProduct.price with its priceLocale.
_NS_NUMBER_FORMATTER_CURRENCY = 2


def error_is_cancelled(error) -> bool:
    """True when a StoreKit NSError means the user cancelled the purchase."""
    try:
        return (
            error is not None
            and str(error.domain()) == _SK_ERROR_DOMAIN
            and int(error.code()) == _SK_ERROR_PAYMENT_CANCELLED
        )
    except Exception:  # noqa: BLE001 - foreign error object; never raise on UI path
        return False

# Defensive import: a source run may not have StoreKit bundled. When it's
# missing we keep the whole API surface working (no-ops + dev-unlocked Pro) so
# the app — and the converter tests that import tablemark.app — still load.
try:
    import StoreKit as _StoreKit  # noqa: N813
except Exception as exc:  # noqa: BLE001 - importing must never crash startup
    _StoreKit = None
    log(f"StoreKit unavailable (dev/non-MAS build?): {exc}")

_SKProductsRequest = getattr(_StoreKit, "SKProductsRequest", None) if _StoreKit else None
_SKPayment = getattr(_StoreKit, "SKPayment", None) if _StoreKit else None
_SKPaymentQueue = getattr(_StoreKit, "SKPaymentQueue", None) if _StoreKit else None


class _ProductsDelegate(NSObject):
    """SKProductsRequestDelegate — receives fetched product metadata.

    PyObjC maps Obj-C selectors to underscore method names. Keep these exact:
    ``productsRequest:didReceiveResponse:`` and ``request:didFailWithError:``.
    """

    def initWithStore_(self, store):
        self = objc_super_init(self)
        if self is None:
            return None
        self._store = store
        return self

    def productsRequest_didReceiveResponse_(self, request, response):
        products = {}
        try:
            for product in response.products():
                products[str(product.productIdentifier())] = product
            for invalid in response.invalidProductIdentifiers() or []:
                log(f"invalid product id: {invalid}")
        except Exception as exc:  # noqa: BLE001
            log(f"parsing products response failed: {exc}")
        self._store._on_products(products)

    def request_didFailWithError_(self, request, error):
        log(f"SKProductsRequest failed: {error}")
        self._store._on_products_failed(error)


class _PaymentObserver(NSObject):
    """SKPaymentTransactionObserver — drives every transaction to completion.

    Every transaction reaching purchased/restored/failed MUST be finished with
    ``finishTransaction_`` or StoreKit re-delivers it forever (and consumables
    can't be re-bought). Selector mapped name:
    ``paymentQueue:updatedTransactions:``.
    """

    def initWithStore_(self, store):
        self = objc_super_init(self)
        if self is None:
            return None
        self._store = store
        return self

    def paymentQueue_updatedTransactions_(self, queue, transactions):
        for txn in transactions:
            try:
                state = int(txn.transactionState())
            except Exception as exc:  # noqa: BLE001
                log(f"reading transaction state failed: {exc}")
                continue

            if state == _STATE_PURCHASING:
                continue  # in flight — wait for a terminal state
            if state in (_STATE_PURCHASED, _STATE_RESTORED):
                self._store._on_transaction_success(
                    txn, restored=(state == _STATE_RESTORED)
                )
                queue.finishTransaction_(txn)
            elif state == _STATE_FAILED:
                self._store._on_transaction_failed(txn)
                queue.finishTransaction_(txn)
            elif state == _STATE_DEFERRED:
                # "Ask to Buy" / parental approval pending — leave it; StoreKit
                # re-delivers with a terminal state once resolved. Do NOT finish.
                log("transaction deferred (awaiting approval)")

    def paymentQueueRestoreCompletedTransactionsFinished_(self, queue):
        log("restore completed")
        self._store._on_restore_finished(None)

    def paymentQueue_restoreCompletedTransactionsFailedWithError_(self, queue, error):
        log(f"restore failed: {error}")
        self._store._on_restore_finished(error)


def objc_super_init(obj):
    """Call NSObject's designated initializer for a PyObjC subclass instance."""
    return obj.init()


class Store:
    """Facade over StoreKit 1 for donations and the XML Pro subscription.

    Public API:
        fetch_products()          -> kick off async product metadata fetch
        purchase_tip(tier)        -> buy a consumable donation ('small'|'medium'|'large')
        subscribe_pro()           -> buy/renew the yearly Pro subscription
        restore()                 -> restore prior purchases (Apple-required button)
        pro_active (property)     -> bool: is the XML Pro entitlement currently valid

    Callbacks (optional, set by the app to refresh UI / show feedback). All may
    fire on a background thread — UI work must hop to the main thread:
        on_pro_changed(active: bool)
        on_tip_purchased(product_id: str)
        on_purchase_failed(product_id: str | None, error)
        on_restore_finished(restored_count: int, error)
        on_products_loaded()
    """

    def __init__(self):
        self.products = {}
        self._products_loaded = False
        self._fetch_failed = False
        self._fetch_retries = 0
        self._restored_count = 0
        self.on_pro_changed = None
        self.on_tip_purchased = None
        self.on_purchase_failed = None
        self.on_restore_finished = None
        self.on_products_loaded = None

        # Strong refs: PyObjC delegates/observers are plain Obj-C objects with no
        # owner on the Python side. Without these attributes they'd be GC'd and
        # StoreKit would call into freed memory. Keep them alive for the app's life.
        self._delegate = None
        self._observer = None
        self._request = None

        self._available = _SKPaymentQueue is not None
        if self._available:
            try:
                self._observer = _PaymentObserver.alloc().initWithStore_(self)
                _SKPaymentQueue.defaultQueue().addTransactionObserver_(self._observer)
            except Exception as exc:  # noqa: BLE001
                log(f"failed to register payment observer: {exc}")
                self._available = False
        else:
            log("StoreKit payment queue unavailable; running dev-unlocked")

    # --- Entitlement state ---

    @property
    def pro_active(self) -> bool:
        """True when the XML Pro entitlement is currently valid.

        MVP verification (spec §3.5): we cache the subscription expiry at
        purchase/renew/restore time and gate on ``now < expiry`` (+ grace). The
        renewal transaction arrives automatically while the app runs and bumps
        the cached expiry forward.

        Dev-only convenience: on a source/non-MAS build where StoreKit isn't
        available we return True so the developer can exercise XML without a
        sandbox account. This branch never runs in a real App Store build (the
        payment queue is always present there).

        TODO(hardening): replace the cached-expiry heuristic with precise
        verification — parse the local ``_MASReceipt/receipt`` (PKCS#7 + ASN.1,
        e.g. bundling ``asn1crypto``) for ``expires_date``, or validate via the
        App Store Server API. The cache can lag reality on long offline spells.
        """
        if not self._available:
            return True  # dev-only: framework absent -> unlocked for development
        if settings.load_pro_active():
            return True
        expires = settings.load_pro_expires()
        return expires > 0 and time.time() < (expires + _GRACE_SECONDS)

    def _set_pro(self, active: bool, expires: float | None = None) -> None:
        settings.save_pro_active(active)
        if expires is not None:
            settings.save_pro_expires(float(expires))
        if callable(self.on_pro_changed):
            try:
                self.on_pro_changed(self.pro_active)
            except Exception as exc:  # noqa: BLE001
                log(f"on_pro_changed callback raised: {exc}")

    # --- Product fetch ---

    def fetch_products(self) -> None:
        """Start an async SKProductsRequest for all known product IDs.

        Results arrive on the delegate; ``self.products`` maps id -> SKProduct
        (carrying localized price/title). Safe no-op when StoreKit is absent.
        """
        if not self._available or _SKProductsRequest is None:
            self._fetch_failed = True
            return
        try:
            ids = set(ALL_PRODUCT_IDS)
            self._fetch_failed = False
            self._delegate = _ProductsDelegate.alloc().initWithStore_(self)
            request = _SKProductsRequest.alloc().initWithProductIdentifiers_(ids)
            request.setDelegate_(self._delegate)
            self._request = request  # strong ref until it completes
            request.start()
            log("SKProductsRequest started")
        except Exception as exc:  # noqa: BLE001
            log(f"fetch_products failed: {exc}")
            self._fetch_failed = True

    def localized_price(self, product_id: str) -> str | None:
        """Localized price string (e.g. '₩6,600', '$4.99') for a fetched product.

        None until product metadata has arrived — callers keep their fallback
        text in that case. Formatting uses the SKProduct's own priceLocale so
        the currency always matches the user's storefront.
        """
        product = self.products.get(product_id)
        if product is None:
            return None
        try:
            formatter = NSNumberFormatter.alloc().init()
            formatter.setNumberStyle_(_NS_NUMBER_FORMATTER_CURRENCY)
            locale = product.priceLocale()
            if locale is not None:
                formatter.setLocale_(locale)
            text = formatter.stringFromNumber_(product.price())
            return str(text) if text else None
        except Exception as exc:  # noqa: BLE001
            log(f"formatting price failed for {product_id}: {exc}")
            return None

    def _on_products(self, products: dict) -> None:
        self.products = products
        self._products_loaded = True
        self._request = None
        log(f"loaded {len(products)} products")
        if callable(self.on_products_loaded):
            try:
                self.on_products_loaded()
            except Exception as exc:  # noqa: BLE001
                log(f"on_products_loaded callback raised: {exc}")

    def _on_products_failed(self, error) -> None:
        self._fetch_failed = True
        self._request = None
        if self._fetch_retries < _MAX_FETCH_RETRIES:
            self._fetch_retries += 1
            delay = _FETCH_RETRY_SECONDS * self._fetch_retries
            log(f"retrying product fetch in {delay}s ({self._fetch_retries}/{_MAX_FETCH_RETRIES})")
            timer = threading.Timer(delay, self.fetch_products)
            timer.daemon = True
            timer.start()

    # --- Purchasing ---

    def purchase_tip(self, tier: str) -> None:
        """Buy a consumable donation. ``tier`` is 'small' | 'medium' | 'large'."""
        product_id = TIP_PRODUCT_IDS.get(tier)
        if product_id is None:
            log(f"unknown tip tier: {tier}")
            return
        self._purchase(product_id)

    def subscribe_pro(self) -> None:
        """Buy (or renew) the yearly XML Pro subscription."""
        self._purchase(PRO_YEARLY)

    def _purchase(self, product_id: str) -> None:
        if not self._available or _SKPayment is None or _SKPaymentQueue is None:
            log(f"cannot purchase {product_id}: StoreKit unavailable")
            self._fail(product_id, None)
            return
        product = self.products.get(product_id)
        try:
            if product is not None:
                payment = _SKPayment.paymentWithProduct_(product)
            else:
                # Products not fetched yet — fall back to building a payment by
                # id so a click isn't dead while metadata is still loading.
                payment = _SKPayment.alloc().init()
                if hasattr(payment, "setProductIdentifier_"):
                    payment.setProductIdentifier_(product_id)
            _SKPaymentQueue.defaultQueue().addPayment_(payment)
            log(f"payment queued: {product_id}")
        except Exception as exc:  # noqa: BLE001
            log(f"queueing payment failed for {product_id}: {exc}")
            self._fail(product_id, None)

    def restore(self) -> None:
        """Restore previous purchases (the Apple-required 'Restore' action)."""
        if not self._available or _SKPaymentQueue is None:
            log("cannot restore: StoreKit unavailable")
            if callable(self.on_restore_finished):
                try:
                    self.on_restore_finished(0, None)
                except Exception as exc:  # noqa: BLE001
                    log(f"on_restore_finished callback raised: {exc}")
            return
        self._restored_count = 0
        try:
            _SKPaymentQueue.defaultQueue().restoreCompletedTransactions()
            log("restore requested")
        except Exception as exc:  # noqa: BLE001
            log(f"restore failed: {exc}")

    # --- Transaction outcome handlers (called by the observer) ---

    def _on_transaction_success(self, txn, restored: bool = False) -> None:
        product_id = self._transaction_product_id(txn)
        log(f"transaction success: {product_id}{' (restored)' if restored else ''}")
        if restored:
            self._restored_count += 1
        if product_id == PRO_YEARLY:
            self._set_pro(True, self._compute_expiry(txn))
        elif product_id in TIP_PRODUCT_IDS.values():
            if callable(self.on_tip_purchased):
                try:
                    self.on_tip_purchased(product_id)
                except Exception as exc:  # noqa: BLE001
                    log(f"on_tip_purchased callback raised: {exc}")

    def _on_transaction_failed(self, txn) -> None:
        product_id = self._transaction_product_id(txn)
        error = None
        try:
            error = txn.error()
        except Exception:  # noqa: BLE001
            pass
        self._fail(product_id, error)

    def _on_restore_finished(self, error) -> None:
        # Restored transactions already flowed through _on_transaction_success
        # (state == restored), so Pro is updated by the time we get here. Report
        # the outcome so the app can confirm the Apple-required Restore action
        # (success / nothing to restore / failure) instead of staying silent.
        count = self._restored_count
        self._restored_count = 0
        if callable(self.on_restore_finished):
            try:
                self.on_restore_finished(count, error)
            except Exception as exc:  # noqa: BLE001
                log(f"on_restore_finished callback raised: {exc}")

    def _fail(self, product_id, error) -> None:
        if callable(self.on_purchase_failed):
            try:
                self.on_purchase_failed(product_id, error)
            except Exception as exc:  # noqa: BLE001
                log(f"on_purchase_failed callback raised: {exc}")

    # --- Helpers ---

    @staticmethod
    def _transaction_product_id(txn):
        try:
            return str(txn.payment().productIdentifier())
        except Exception:  # noqa: BLE001
            return None

    @staticmethod
    def _compute_expiry(txn) -> float:
        """Best-effort subscription expiry as an epoch timestamp.

        StoreKit 1's SKPaymentTransaction does not expose the expiry directly, so
        the MVP derives it from the transaction date + one year. Precise expiry
        needs receipt parsing — see the TODO on ``pro_active``.
        """
        base = time.time()
        try:
            date = txn.transactionDate()
            if date is not None:
                base = float(date.timeIntervalSince1970())
        except Exception:  # noqa: BLE001
            pass
        return base + _YEAR_SECONDS
