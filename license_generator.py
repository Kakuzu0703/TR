import json
import hashlib
import os
import datetime
import argparse
import base64
import hmac
import secrets
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC


class LicenseGenerator:
    def __init__(self, secret_key=None):
        """
        初始化许可证生成器
        :param secret_key: 用于签名的密钥，如果不提供则自动生成
        """
        # 如果没有提供密钥，则生成一个
        if secret_key is None:
            self.secret_key = secrets.token_hex(32)  # 生成64字符的十六进制字符串
        else:
            self.secret_key = secret_key

        # 保存密钥到密钥文件（在实际应用中，这个文件应该被安全保存）
        with open('license_secret.key', 'w') as f:
            f.write(self.secret_key)

        print(f"许可证密钥已保存到 license_secret.key")

    def load_fingerprint(self, fingerprint_file):
        """
        加载硬件指纹文件
        :param fingerprint_file: 硬件指纹文件路径
        :return: 硬件指纹数据
        """
        try:
            with open(fingerprint_file, 'r') as f:
                fingerprint_data = json.load(f)
            return fingerprint_data
        except Exception as e:
            print(f"加载指纹文件失败: {e}")
            return None

    def create_license(self, fingerprint_data, expiry_days=365, user_info=None):
        """
        创建许可证数据
        :param fingerprint_data: 硬件指纹数据
        :param expiry_days: 许可证有效期（天）
        :param user_info: 用户信息字典
        :return: 许可证数据
        """
        # 获取硬件唯一ID
        unique_id = fingerprint_data.get('unique_id')
        if not unique_id:
            print("错误: 指纹数据中不包含唯一ID")
            return None

        # 创建许可证数据
        issue_date = datetime.datetime.now()
        expiry_date = issue_date + datetime.timedelta(days=expiry_days)

        license_data = {
            'hardware_id': unique_id,
            'issue_date': issue_date.strftime('%Y-%m-%d %H:%M:%S'),
            'expiry_date': expiry_date.strftime('%Y-%m-%d %H:%M:%S'),
            'allowed_features': ['base', 'premium', 'export'],  # 可自定义的功能列表
        }

        # 如果提供了用户信息，添加到许可证
        if user_info:
            license_data['user'] = user_info

        # 创建许可证签名
        signature = self._sign_data(license_data)
        license_data['signature'] = signature

        return license_data

    def generate_license_file(self, license_data, output_file='license.dat'):
        """
        生成加密的许可证文件
        :param license_data: 许可证数据
        :param output_file: 输出文件路径
        """
        # 将许可证数据转换为JSON字符串
        license_json = json.dumps(license_data)

        # 生成加密密钥
        salt = os.urandom(16)
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(self.secret_key.encode()))

        # 加密许可证数据
        f = Fernet(key)
        encrypted_license = f.encrypt(license_json.encode())

        # 将salt添加到加密数据前面，以便解密时使用
        final_data = salt + encrypted_license

        # 保存到文件
        with open(output_file, 'wb') as f:
            f.write(final_data)

        print(f"已生成许可证文件: {output_file}")

        # 同时生成一个用于参考的可读文本文件
        readable_file = f"{os.path.splitext(output_file)[0]}_info.txt"
        with open(readable_file, 'w') as f:
            f.write("许可证信息（仅供参考）:\n\n")
            for key, value in license_data.items():
                if key != 'signature':  # 签名太长，不显示
                    f.write(f"{key}: {value}\n")
                else:
                    f.write(f"{key}: [签名数据...]\n")

        print(f"已生成参考文件: {readable_file}")

    def _sign_data(self, data):
        """
        为数据创建签名
        :param data: 要签名的数据
        :return: 签名字符串
        """
        # 创建数据的一个副本，移除任何现有签名
        data_copy = data.copy()
        data_copy.pop('signature', None)

        # 对数据进行排序，确保一致性
        data_str = json.dumps(data_copy, sort_keys=True)

        # 创建HMAC签名
        signature = hmac.new(
            self.secret_key.encode(),
            data_str.encode(),
            hashlib.sha256
        ).hexdigest()

        return signature


def main():
    parser = argparse.ArgumentParser(description='生成软件许可证')
    parser.add_argument('--fingerprint', '-f', required=True, help='硬件指纹JSON文件路径')
    parser.add_argument('--output', '-o', default='license.dat', help='输出的许可证文件路径')
    parser.add_argument('--days', '-d', type=int, default=365, help='许可证有效期（天）')
    parser.add_argument('--name', '-n', help='用户/公司名称')
    parser.add_argument('--email', '-e', help='用户/公司电子邮件')
    parser.add_argument('--key', '-k', help='许可证密钥（如不提供则自动生成）')

    args = parser.parse_args()

    # 创建许可证生成器
    generator = LicenseGenerator(args.key)

    # 加载硬件指纹
    fingerprint_data = generator.load_fingerprint(args.fingerprint)
    if not fingerprint_data:
        return

    # 准备用户信息
    user_info = None
    if args.name or args.email:
        user_info = {}
        if args.name:
            user_info['name'] = args.name
        if args.email:
            user_info['email'] = args.email

    # 创建许可证
    license_data = generator.create_license(fingerprint_data, args.days, user_info)
    if license_data:
        # 生成许可证文件
        generator.generate_license_file(license_data, args.output)


if __name__ == "__main__":
    main()


# # 生成许可证（自动创建密钥）
# python license_generator.py -f hardware_fingerprint.json -d 365 -n "测试用户" -o license.dat
#
# # 使用已有密钥
# python license_generator.py -f fingerprint.json -k existing_key -o custom.dat