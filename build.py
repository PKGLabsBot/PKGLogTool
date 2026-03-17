import os
import shutil
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

VERSION_FILE = BASE_DIR / "version.txt"
ICON_FILE = BASE_DIR / "icon.ico"
VERSION_INFO = BASE_DIR / "file_version_info.txt"
MAIN_SCRIPT = BASE_DIR / "logtool.py"

BUILD_DIR = BASE_DIR / "build"
DIST_DIR = BASE_DIR / "dist"
RELEASE_DIR = BASE_DIR / "release"

# ============================================================
# PyInstaller에서 자동 탐지되지 않는 동적 import 목록
# pandas가 엑셀 엔진을 importlib으로 로드하기 때문에
# xlsxwriter, openpyxl 등을 명시적으로 포함해야 함
# ============================================================
HIDDEN_IMPORTS = [
    "xlsxwriter",
    "xlsxwriter.workbook",
    "xlsxwriter.worksheet",
    "xlsxwriter.chart",
    "xlsxwriter.utility",
    "pandas.io.excel._xlsxwriter",
    "pandas.io.excel._openpyxl",
    "pandas.io.formats.excel",
]

# 제외할 불필요한 대형 패키지 (exe 크기 축소)
EXCLUDES = [
    "matplotlib",
    "scipy",
    "PIL",
    "IPython",
    "notebook",
    "pytest",
    "setuptools",
    "pkg_resources",
]


def read_version():
    if not VERSION_FILE.exists():
        return "1.0"
    return VERSION_FILE.read_text(encoding="utf-8").strip()


def bump_version(version):
    parts = version.split(".")
    if len(parts) < 2:
        return f"{version}.1"
    major = parts[0]
    minor = int(parts[1]) + 1
    return f"{major}.{minor}"


def write_version(version):
    VERSION_FILE.write_text(version, encoding="utf-8")


def clean():
    for d in [BUILD_DIR, DIST_DIR]:
        if d.exists():
            shutil.rmtree(d)
            print(f"  cleaned: {d}")


def build(version):
    exe_name = f"LogTool_v{version}"

    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        f"--icon={ICON_FILE}",
        f"--name={exe_name}",
        f"--version-file={VERSION_INFO}",
    ]

    # hidden imports 추가
    for hi in HIDDEN_IMPORTS:
        cmd += ["--hidden-import", hi]

    # 불필요 패키지 제외 (exe 크기 축소)
    for ex in EXCLUDES:
        cmd += ["--exclude-module", ex]

    cmd.append(str(MAIN_SCRIPT))

    print(f"\n[BUILD] pyinstaller 실행 중... (exe 이름: {exe_name})")
    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode != 0:
        print("\n[ERROR] 빌드 실패. 위 에러 메시지를 확인하세요.")
        raise RuntimeError("pyinstaller build failed")

    exe_path = DIST_DIR / f"{exe_name}.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"빌드 결과 exe를 찾을 수 없습니다: {exe_path}")

    print(f"\n[BUILD] 빌드 성공: {exe_path}")
    return exe_path


def create_release(version, exe_path: Path):
    exe_name = f"LogTool_v{version}"
    release_folder = RELEASE_DIR / exe_name

    if release_folder.exists():
        shutil.rmtree(release_folder)
    release_folder.mkdir(parents=True)

    # exe 복사
    shutil.copy2(exe_path, release_folder / f"{exe_name}.exe")

    # profiles 폴더 복사 (필수 - 설비군 설정)
    if (BASE_DIR / "profiles").exists():
        shutil.copytree(BASE_DIR / "profiles", release_folder / "profiles")

    # version.txt 복사
    if VERSION_FILE.exists():
        shutil.copy2(VERSION_FILE, release_folder / "version.txt")

    # README 생성
    readme = f"""LogTool v{version} 사용 방법
=================================

[사전 설치 불필요 - 단독 실행 가능]

1. {exe_name}.exe 실행
2. 설비군 선택 (profiles 폴더의 JSON 파일이 자동 적용됨)
3. Handler 로그 폴더 선택
4. Vision 로그 폴더 선택
5. 리포트 생성 클릭
6. output 폴더에 Excel 파일 저장됨

[profiles 커스터마이징]
- profiles 폴더 내 JSON 파일을 수정하면 설비군별 파서 설정 변경 가능
- 새 설비군 추가 시 JSON 파일 추가 후 재실행

[문제 발생 시]
- debug_log.txt 파일 확인
- debug_samples.txt 에서 파싱 실패 라인 확인

버전: {version}
"""
    (release_folder / "README.txt").write_text(readme.strip(), encoding="utf-8")

    print(f"\n[RELEASE] 배포 폴더 생성 완료: {release_folder}")
    return release_folder


def main():
    print("=" * 50)
    print("LogTool 빌드 스크립트")
    print("=" * 50)

    old_version = read_version()
    new_version = bump_version(old_version)

    print(f"  이전 버전: {old_version}")
    print(f"  새 버전:   {new_version}")

    answer = input("\n버전을 올리고 빌드하시겠습니까? (y/n): ").strip().lower()
    if answer != "y":
        print("빌드 취소.")
        return

    write_version(new_version)

    clean()
    exe_path = build(new_version)
    release_folder = create_release(new_version, exe_path)

    print("\n" + "=" * 50)
    print("빌드 완료")
    print(f"배포 폴더: {release_folder}")
    print("=" * 50)


if __name__ == "__main__":
    main()
