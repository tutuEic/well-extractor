"""
Key Generator - 输入机器码生成授权密钥
"""
import hashlib
import sys

SALT = "HYDRA-WELL-2026-X8K3M"

def generate_key(machine_code):
    clean = machine_code.replace('-', '').strip().upper()
    h = hashlib.sha256((clean + SALT).encode()).hexdigest()[:16].upper()
    return '-'.join([h[i:i+4] for i in range(0, 16, 4)])

if __name__ == '__main__':
    if len(sys.argv) > 1:
        mc = sys.argv[1]
    else:
        mc = input('请输入机器码: ')
    key = generate_key(mc)
    print(f'\n机器码: {mc}')
    print(f'授权密钥: {key}')
    input('\n按回车退出...')
