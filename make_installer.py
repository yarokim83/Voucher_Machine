"""Build application and setup from checked-in specifications."""
import subprocess
import sys
from pathlib import Path


def main():
    root = Path(__file__).resolve().parent
    version = (root / 'version.txt').read_text(encoding='utf-8').strip()
    import re
    if not re.fullmatch(r'v\d+\.\d+\.\d+', version):
        raise ValueError('Invalid version.txt')
    for spec in ('VoucherPass.spec', 'VoucherPass_Setup.spec'):
        subprocess.run([sys.executable, '-m', 'PyInstaller', '--clean', '--noconfirm',
                        '--distpath', str(root / 'dist' / version),
                        '--workpath', str(root / 'build' / version), str(root / spec)], cwd=root, check=True)


if __name__ == '__main__':
    main()
