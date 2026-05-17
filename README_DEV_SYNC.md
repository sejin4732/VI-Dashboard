# VICT 개발/동기화 안내

## 아주 짧게

개인PC:
- `git add .`
- `git commit -m "수정내용"`
- `git push`

클라우드PC:
- `git pull`

## 기본 원칙

- Git에는 코드만 올립니다.
- `parquet`, `Excel`, `cache`, `outputs`, `logs`, `bom_lake`, 배포 폴더는 Git에 올리지 않습니다.
- 클라우드 PC 데이터는 `D:/VICT_DATA` 아래에 유지합니다.
- 개인 PC에서는 코드 수정 후 `git push`만 하면 됩니다.
- 클라우드 PC에서는 `git pull` 후 바로 테스트하면 됩니다.

## 개인 PC에서 수정 후 올리기

1. `git status`
2. `git add .`
3. `git commit -m "수정내용"`
4. `git push`

최초 1회는 아래 스크립트로 Git 초기 구성을 할 수 있습니다.

```powershell
.\setup_git_personal_pc.ps1
```

## 클라우드 PC에서 최신 코드 받기

1. 아래 스크립트 실행

```powershell
.\update_cloud_pc.ps1
```

2. 그 다음 GUI 실행

```powershell
python vi_report_gui.py
```

## 클라우드 PC 로컬 데이터 경로

- `D:/VICT_DATA/data`
- `D:/VICT_DATA/bom_lake`
- `D:/VICT_DATA/cache`
- `D:/VICT_DATA/outputs`
- `D:/VICT_DATA/logs`

`update_cloud_pc.ps1`가 `config.json`이 없으면 자동으로 만들어 줍니다.
