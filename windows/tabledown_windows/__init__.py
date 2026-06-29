"""Windows port of Tabledown."""

# Windows 포트가 자기 버전을 소유한다. Windows 스토어(MSIX)는 이 값을 읽어
# build_msix.ps1 이 PackageVersion 을 만든다. 공유 tablemark.__version__(현재
# 0.4.0)은 아직 테스트 중인 다음 macOS 버전이라 일부러 따르지 않는다 — 그걸 쓰면
# MSIX 가 0.4.0.0 으로 잘못 찍힌다.
# 0.2.5: '로그인 시 자동 실행' 토글 + 도움말 창 안 닫힘 수정 (0.2.4 출시 이후).
# 0.2.6: 단일 인스턴스 가드 — 트레이/클립보드 워처 중복 실행 방지(macOS 는 LaunchServices 가 기본 제공).
__version__ = "0.2.6"
