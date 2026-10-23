티끌 알림봇 V8.30 통합 수정본

포함 파일:
- monitor.py
- .github/workflows/tikkle_monitor.yml

GitHub 저장소에서 반드시 같은 경로의 기존 파일을 교체하세요.
1. 저장소 루트의 monitor.py 교체
2. .github/workflows/tikkle_monitor.yml 교체
3. monitor_patch.py는 더 이상 워크플로에서 실행하지 않습니다. 기존 파일은 남겨둬도 됩니다.
4. 두 파일을 커밋한 후 Actions에서 수동 실행하여 로그를 확인하세요.

사전 확인:
- Python 컴파일 통과
- 워크플로 설정 검증에 해당하는 15개 assertion 통과
- 실제 GitHub Actions 실행은 아직 확인되지 않았습니다.
