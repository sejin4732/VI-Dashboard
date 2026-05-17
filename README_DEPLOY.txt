AX_Deploy\VICT 는 사내클라우드의 \\...\DX\VICT 폴더에 그대로 복붙하기 위한 배포용 묶음입니다.

상위 실행 파일
- vi_report_gui.py
- streamlit_vi_dashboard.py
- vi_report_cli.py
- vi_calculation_engine.py
- bom_lake.py
- historical_bom_lake.py
- requirements.txt

역할별 폴더 구조
- workflow\ : 추천/집계/BOM 관련 실제 구현
- dashboard\ : Streamlit 대시보드 화면 구현
- bom_lake\ : parquet 데이터 저장소
- compat_imports\ : 호환용 import 패키지
- VI Result\ : 결과 파일 보관용 폴더

세부 폴더
- workflow\
- dashboard\core
- dashboard\assets
- bom_lake\historical
- bom_lake\current_cache
- bom_lake\metadata
- VI Result

배포 방법
1. 사내클라우드의 기존 VICT 폴더를 백업합니다.
2. AX_Deploy\VICT 폴더 안의 내용으로 기존 VICT 폴더를 덮어씁니다.
3. 기존에 쓰던 parquet 파일이 있으면 bom_lake 폴더 아래로 옮깁니다.
4. GUI는 반드시 덮어쓴 VICT 폴더 안의 vi_report_gui.py 로 실행합니다.

주의
- AX 로컬 폴더를 수정해도, 사내클라우드의 VICT 복사본을 덮어쓰지 않으면 화면은 바뀌지 않습니다.
- dashboard_launch.log 에 찍히는 script 경로가 현재 실행 중인 실제 복사본입니다.
