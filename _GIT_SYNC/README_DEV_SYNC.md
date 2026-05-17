# VICT Git 동기화 안내

## 아주 짧게

개인PC:
- `git add .`
- `git commit -m "수정내용"`
- `git push`

클라우드PC:
- `git pull`

## 파일 위치

- Git 관련 문서/스크립트는 `_GIT_SYNC` 폴더에서 관리합니다.
- 단, `.gitignore`는 Git이 루트에서 읽어야 하므로 프로젝트 루트에 그대로 둡니다.

## 기본 원칙

- Git에는 코드만 올립니다.
- `parquet`, `Excel`, `cache`, `outputs`, `logs`, `bom_lake`, 배포 폴더는 Git에 올리지 않습니다.
- 클라우드 PC 데이터는 `D:/VICT_DATA` 아래에 유지합니다.

## 개인 PC에서 최초 1회

```powershell
.\_GIT_SYNC\setup_git_personal_pc.ps1
```

## 클라우드 PC에서 최신 코드 받기

```powershell
.\_GIT_SYNC\update_cloud_pc.ps1
python vi_report_gui.py
```

## 클라우드 PC 로컬 데이터 경로

- `D:/VICT_DATA/data`
- `D:/VICT_DATA/bom_lake`
- `D:/VICT_DATA/cache`
- `D:/VICT_DATA/outputs`
- `D:/VICT_DATA/logs`

`_GIT_SYNC\update_cloud_pc.ps1`가 `config.json`이 없으면 자동으로 만들어 줍니다.
