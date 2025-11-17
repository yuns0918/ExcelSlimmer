# ExcelSlimmer

ExcelSlimmer는 다음 세 가지 Excel 최적화 도구를 **한 번에 파이프라인으로 실행**할 수 있게 해주는 통합 런처입니다.
메인 UI는 PySide6(Qt) 기반의 `excel_slimmer_qt.py`로 구현되어 있습니다.

1. **ExcelCleaner** – 이름 정의(definedNames) 정리
2. **Excel Image Slimmer** – 이미지 리사이즈/압축
3. **Excel Slimmer – Precision Plus** – 고급 용량 감소(정밀 슬리머)

사용자는 단일 UI에서 파일을 선택하고, 사용할 기능만 체크한 뒤 순서대로 실행할 수 있습니다.

---

## 0. 사내 배포/사용 요약

- 일반 사용자(사내 직원)에게는 다음 두 파일을 **ZIP으로 묶어 배포**하는 것을 권장합니다.
  - `ExcelSlimmer.exe` – 실행 파일 (PyInstaller onefile)
  - `README.txt` – 직원용 간단 사용 설명서
- 직원 측 사용 흐름
  1. ZIP 압축 해제
  2. `ExcelSlimmer.exe` 더블 클릭
  3. 상세 사용법 및 주의 사항은 같은 폴더의 `README.txt` 참고

> 이 저장소의 `README.md`는 **개발/빌드 관점의 문서**이며,
> 실제 배포 시에는 `ExcelSlimmer.exe` + `README.txt` 조합만 있으면 충분합니다.

---

## 1. 폴더/파일 구성 (EXE/데스크톱용)

이 저장소(ExcelSlimmerEXE)는 **단일 폴더 안에 EXE 빌드에 필요한 모든 코드**를 포함합니다.

- `excel_slimmer_qt.py`  
  PySide6(Qt) 기반 메인 데스크톱 UI
- `excel_suite_pipeline.py`  
  이름 정리/이미지 최적화/정밀 슬리머를 묶는 공통 파이프라인 로직 (데스크톱·웹 공용)
- `backData/`  
  기존 Tk 기반 도구 코드(ExcelCleaner, Image Slimmer, Precision Plus)를 모아 둔 폴더
- `settings.py`  
  설정 저장/로드 (테마, 출력 폴더, 로그 옵션 등)
- `install.bat` / `run.bat` / `build.bat`  
  각각 개발 환경 준비, Qt UI 실행, 단일 EXE 빌드 스크립트
- `ExcelSlimmer.spec` / `ExcelSlimmer.ico` / `check_white.svg`  
  PyInstaller 빌드 설정, EXE 아이콘, 체크박스 체크 아이콘

> 예전에는 `ExcelCleaner`, `ExcelImageOptimization`, `ExcelByteReduce`가 별도 폴더로 존재했지만,
> 현재 EXE 빌드에는 **이 저장소 하나(`ExcelSlimmerEXE`)**만 있으면 됩니다.

> 웹 버전 코드는 별도 저장소 **ExcelSlimmerWeb**에서 관리합니다.

---

## 2. 개발 환경 준비 (install.bat)

1. Windows에서 `ExcelSuite` 폴더로 이동
2. `install.bat` 더블 클릭 실행
3. 스크립트가 자동으로 수행하는 작업
   - `.venv_suite` 가상환경 생성 (없으면 생성, 있으면 재사용)
   - `pip / wheel / setuptools` 업그레이드
   - 필수 패키지 설치
     - `pillow`
     - `lxml`
     - `pyinstaller`
4. `[OK] Environment ready.` 메시지가 나오면 준비 완료

---

## 3. 통합 도구 실행 (run.bat)

1. `ExcelSuite` 폴더에서 `run.bat` 실행
2. `.venv_suite` 가상환경이 활성화되고, PySide6 기반 UI `excel_slimmer_qt.py`가 실행됩니다.
3. UI에서 다음과 같이 사용합니다.
   - **[파일 찾기]** 버튼으로 대상 파일 선택 (`.xlsx` / `.xlsm`)
   - **실행할 기능** 체크
     - `1) 이름 정의 정리`
     - `2) 이미지 최적화`
     - `3) 정밀 슬리머 (Precision Plus)`
   - **정밀 슬리머 옵션** (정밀 슬리머를 켠 경우에만 활성화)
     - 공격 모드 (이미지 리사이즈 + PNG→JPG 변환)
     - XML 정리 (calcChain, printerSettings 등 구조 정리)
     - 숨은 XML 데이터 삭제 (customXml) — *주의, 일반적으로 비권장*
   - **[선택한 기능 실행]** 버튼 클릭
4. 하단 로그/상태바를 통해 진행 상황을 확인할 수 있습니다.

> 정밀 슬리머 사용 시, 일부 파일에서는 Excel이 파일을 열 때
> "이 통합 문서의 내용을 복구하시겠습니까?"라는 안내를 표시할 수 있습니다.
> 이는 Excel이 내부 구조/캐시 변화를 감지했을 때 보여주는 일반적인 동작입니다.

---

## 4. 배포용 단일 EXE 빌드 (build.bat)

1. `ExcelSuite` 폴더에서 `build.bat` 실행
2. 스크립트가 수행하는 작업
   - 이전 `build/`, `dist/` 폴더 정리
   - PyInstaller를 사용해 **단일 EXE** 생성
   - EXE 이름: `dist\ExcelSlimmer.exe`
   - 아이콘
     - `ExcelSuite\ExcelSlimmer.ico` 가 존재하면 이를 아이콘으로 사용
     - 없으면 PyInstaller 기본 아이콘 사용
3. `[OK] Build complete.` 메시지가 나오면 빌드 성공입니다.

> **배포 시:** `dist\ExcelSlimmer.exe` 파일 **한 개만** 전달하면 됩니다.
> 상대방 PC에는 Python이 설치되어 있지 않아도 됩니다.

---

## 5. 출력 구조 및 백업 파일

### 5.1 이름 정의 정리 (ExcelCleaner)

- 최상위 출력 폴더: `C:\Users\사용자명\Desktop\ExcelSlimmed\`
- 실행할 때마다 타임스탬프 하위 폴더 생성:

  ```text
  Desktop\ExcelSlimmed\YYYY-MM-DD-HH-MM-SS\
      원본이름_backup.xlsx
      원본이름_clean.xlsx
  ```

- `*_backup.xlsx` : 원본 전체 백업
- `*_clean.xlsx`  : definedNames가 정리된 버전

### 5.2 파이프라인 전체 실행 시

파이프라인(이름 정리 → 이미지 최적화 → 정밀 슬리머)을 모두 켜고 실행하면:

- 중간 산출물(`.clean.xlsx`, `_slim.xlsx` 등)은 **파이프라인이 성공적으로 끝난 후 자동 삭제**
- 타임스탬프 폴더에는 보통 다음과 같이 남습니다.

```text
Desktop\ExcelSlimmed\YYYY-MM-DD-HH-MM-SS\
    원본이름_backup.xlsx      # 이름 정리 직전 원본 백업
    최종파일명_slimmed.xlsx   # 파이프라인 최종 결과
```

### 5.3 정밀 슬리머 단독 사용 시

파이프라인에서 이름 정리를 끄고 **정밀 슬리머만 사용**하는 경우:

- 별도 타임스탬프 폴더가 아니라, **원본 파일이 있던 폴더**에 다음 파일들이 생깁니다.

```text
원본이름_backup.xlsx
원본이름_slimmed.xlsx
```

---

## 6. 정밀 슬리머 사용 시 주의 사항

정밀 슬리머(Precision Plus)는 이미지/구조 최적화를 강하게 수행하는 기능입니다.

- **이미지 최적화**
  - JPEG/PNG 이미지 리사이즈 및 재압축
  - 공격 모드일 경우 PNG→JPG 변환(알파 채널 없는 경우)
- **XML 정리 옵션 ON일 때만**
  - `calcChain.xml` 제거 (Excel이 다시 생성)
  - `xl/printerSettings/*.bin` 제거
  - `docProps/thumbnail.jpeg` 제거
  - `docProps/custom.xml` 제거
- **숨은 XML 데이터 삭제(customXml)**
  - `xl/customXml` 폴더를 통째로 삭제
  - 특정 솔루션/애드인이 사용하는 메타데이터가 포함될 수 있어
    **일반적인 경우 사용을 권장하지 않습니다.**

### 6.1 Excel 복구 안내 팝업

일부 파일에서는 정밀 슬리머 적용 후 Excel을 열 때 다음과 같은 안내가 뜰 수 있습니다.

> "이 통합 문서의 내용에 문제가 있습니다. 이 통합 문서의 내용을 최대한 복구하시겠습니까?"

이는 다음과 같은 이유로 나타날 수 있습니다.

- 이미지/구조 최적화로 인해 Excel이 내부 구조를 다시 계산하려고 할 때
- 캐시/체인/썸네일 등이 제거되었다가 Excel에 의해 재생성될 때

파일이 실제로 손상된 것이 아니라, Excel이 보수적으로 동작하는 경우도 많습니다.

---

## 7. 보안/안전 관련 메모

- 이 도구가 수행하는 작업
  - 선택한 Excel 파일을 임시 폴더에 압축 해제
  - 이미지 파일 리사이즈/압축
  - 선택된 옵션에 따라 일부 XML/리소스 정리
  - 결과 파일을 **새 이름**으로 저장 (원본은 `_backup` 파일로 보존)
- 네트워크 통신, 자기 복제, 레지스트리 조작 등
  악성코드로 오해받을 만한 동작은 포함되어 있지 않습니다.
- 최종 EXE는 PyInstaller가 Python 인터프리터와 필요한
  라이브러리(PySide6, `pillow`, `lxml` 등)를 한 파일로 묶은 것입니다.

---

## 8. 개발 팁

- ExcelCleaner / ImageOptimization / ByteReduce 각 폴더 내 스크립트를 수정하면,
  통합 도구도 그 로직을 그대로 사용합니다.
  - 단, **함수 이름이나 시그니처를 크게 바꾸는 경우**에는
    `excel_suite_pipeline.py` 쪽 호출 코드도 함께 수정해야 합니다.
- GitHub 저장소(데스크톱/EXE): `https://github.com/yuns0918/ExcelSlimmerEXE`
- GitHub 저장소(웹): `https://github.com/yuns0918/ExcelSlimmerWeb`
  - 각 저장소에서 변경 후에는 `git add . && git commit && git push` 로 버전 관리를 유지합니다.
