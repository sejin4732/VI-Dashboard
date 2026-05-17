# VICT 개발/동기화 안내

이 프로젝트는 **코드만 Git으로 관리**하고, 큰 데이터/캐시/출력은 각 PC 로컬에 따로 보관하는 방식으로 사용합니다.

## 기본 원칙

- `parquet`, `Excel`, `cache`, `outputs`, `logs` 는 Git에 올리지 않습니다.
- 클라우드 PC 데이터는 `D:/VICT_DATA` 아래에 계속 유지합니다.
- 개인 PC에서는 코드만 수정해서 `git push` 합니다.
- 클라우드 PC에서는 최신 코드만 `git pull` 해서 바로 테스트합니다.

## 개인 PC에서 수정 후 올리기

1. `git status`
2. `git add .`
3. `git commit -m "수정내용"`
4. `git push`

처음 한 번은 아래 스크립트를 실행해 초기화하면 됩니다.

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

## 데이터는 어디에 보관하나

클라우드 PC에서는 아래 경로를 사용합니다.

- `D:/VICT_DATA/data`
- `D:/VICT_DATA/bom_lake`
- `D:/VICT_DATA/cache`
- `D:/VICT_DATA/outputs`
- `D:/VICT_DATA/logs`

`update_cloud_pc.ps1`가 `config.json`이 없으면 자동으로 만들어 줍니다.

## 중요한 점

- 코드만 Git으로 업데이트합니다.
- 데이터는 Git에 올리지 않습니다.
- 클라우드 PC의 `D:/VICT_DATA` 데이터는 그대로 유지됩니다.
- 그래서 매번 전체 폴더를 압축/복사/해제할 필요가 없습니다.
