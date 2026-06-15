"""Windows port of Tabledown."""

# Windows 포트가 자기 버전을 소유한다. Windows 스토어(MSIX)는 출시된 0.2.4 기능
# 상태로 배포하므로 build_msix.ps1 이 이 값을 읽어 PackageVersion 을 만든다.
# 공유 tablemark.__version__(현재 0.3.0)은 아직 테스트 중인 다음 macOS 버전이라
# 일부러 따르지 않는다 — 그걸 쓰면 MSIX 가 0.3.0.0 으로 잘못 찍힌다.
__version__ = "0.2.4"
