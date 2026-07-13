# Tabledown — Privacy Policy / 개인정보 처리방침

_Last updated: 2026-07-12_

Tabledown is a Windows tray utility that converts spreadsheet tables (Excel /
Google Sheets) to Markdown — and Markdown tables back into a
spreadsheet-pasteable format — by watching the system clipboard.

## English

**Tabledown does not collect, store, sell, or share any personal data.**

- The app reads the current contents of the Windows clipboard locally and, when
  it detects a table, writes the converted text/HTML formats back to the **same
  clipboard**. All conversion happens **on your PC**; nothing is sent to any
  external server.
- When you explicitly choose **Copy table with formulas as XML**, the app locally
  reads cell values, blanks, formula text, and addresses from the current
  Microsoft Excel selection and writes XML to the same clipboard. Cell values
  and formulas are never written to the diagnostic log or sent to any server.
- Tabledown has **no network connections, no telemetry, no analytics, no ads, no
  account sign-in, and no access** to your location, contacts, photos, or files.
- A local diagnostic log is written only to
  `%LOCALAPPDATA%\Tabledown\Tabledown.log` to help confirm the app is working.
  It records short status messages (e.g. "clipboard formats updated") — **not**
  the full clipboard contents — never leaves your device, and you can delete it
  at any time.
- User settings (language, first-run flag, auto-fill blank cells preference) are
  stored locally in `%APPDATA%\Tabledown\settings.json`.

Because clipboard conversion must run in the background, the app uses the
`runFullTrust` capability and the Win32 clipboard APIs. This is required for
passive background conversion; the UWP clipboard API only permits clipboard
access while the app window has focus. No clipboard data is transmitted off the
device under any circumstance.

Contact: <sukmack@gmail.com>

## 한국어

**Tabledown은 사용자의 개인정보를 수집·저장·판매·공유하지 않습니다.**

- 앱은 Windows 클립보드의 현재 내용을 로컬에서 읽고, 표를 감지하면 변환한
  text/HTML 형식을 **같은 클립보드**에 다시 기록합니다. 모든 변환은 **사용자의 PC
  안에서만** 처리되며 외부 서버로 전송되지 않습니다.
- 사용자가 **‘표의 수식을 포함해 XML로 복사’** 를 명시적으로 실행한 경우에만 현재 Microsoft
  Excel 선택 영역의 셀 값·빈칸·수식·주소를 로컬에서 읽어 같은 클립보드에 XML을
  기록합니다. 셀 값과 수식 내용은 진단 로그에 기록되거나 외부 서버로 전송되지 않습니다.
- Tabledown은 **네트워크 연결, 텔레메트리, 분석, 광고, 계정 로그인이 전혀 없으며**,
  위치·연락처·사진·파일에 접근하지 않습니다.
- 진단 로그는 앱 동작 확인용으로 `%LOCALAPPDATA%\Tabledown\Tabledown.log` 에만
  저장됩니다. 짧은 상태 메시지(예: "clipboard formats updated")만 기록하고 클립보드
  원문 전체는 기록하지 않으며, 외부로 전송되지 않고 사용자가 직접 삭제할 수 있습니다.
- 사용자 설정(언어, 첫 실행 플래그, 빈칸 자동 채우기 설정)은
  `%APPDATA%\Tabledown\settings.json` 에 로컬로 저장됩니다.

클립보드 변환은 백그라운드에서 동작해야 하므로 앱은 `runFullTrust` 권한과 Win32
클립보드 API를 사용합니다. UWP 클립보드 API는 앱 창이 포커스를 가졌을 때만 접근이
가능해 백그라운드 상시 변환이 불가능하기 때문입니다. 어떤 경우에도 클립보드 데이터가
기기 밖으로 전송되지 않습니다.

문의: <sukmack@gmail.com>
