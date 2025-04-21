import json
import hashlib
import hmac
import datetime
import logging
import os
import base64
import sys

from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

from hardware_fingerprint import HardwareFingerprint
from PyQt5.QtCore import QObject


class LicenseValidator(QObject):
    # heartbeat_failed = pyqtSignal(str)  # 定义信号
    def __init__(self, license_path='license.dat', key_path=None):
        super().__init__()
        """
        初始化许可证验证器
        :param license_path: 许可证文件路径
        :param key_path: 密钥文件路径（生产环境中应嵌入到代码中）
        """
        self.license_path = license_path
        self.key_path = key_path
        self.license_data = None

    def _load_key(self):
        """
        从文件加载密钥或使用硬编码密钥
        生产环境中应该使用硬编码或混淆的密钥，而不是外部文件
        """
        if self.key_path and os.path.exists(self.key_path):
            try:
                with open(self.key_path, 'r') as f:
                    return f.read().strip()
            except:
                pass

        # 如果没有找到密钥文件，使用硬编码的密钥
        # 注意: 在实际应用中，这个密钥应该是嵌入和混淆的，而不是像这样明文显示
        return "YOUR_HARDCODED_SECRET_KEY_HERE"  # 请替换为您的实际密钥

    def load_license(self):
        """
        加载并解密许可证文件
        :return: 成功返回True，失败返回False
        """
        if not os.path.exists(self.license_path):
            print(f"错误: 找不到许可证文件 {self.license_path}")
            return False

        try:
            # 读取许可证文件
            with open(self.license_path, 'rb') as f:
                encrypted_data = f.read()

            # 前16字节是salt
            salt = encrypted_data[:16]
            encrypted_license = encrypted_data[16:]

            # 获取密钥
            secret_key = self._load_key()

            # 派生加密密钥
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=salt,
                iterations=100000,
            )
            key = base64.urlsafe_b64encode(kdf.derive(secret_key.encode()))

            # 解密
            f = Fernet(key)
            decrypted_data = f.decrypt(encrypted_license)

            # 解析JSON
            self.license_data = json.loads(decrypted_data.decode())
            return True

        except InvalidToken:
            print("错误: 许可证文件无效或已损坏")
            return False
        except Exception as e:
            print(f"解析许可证时发生错误: {e}")
            return False

    def _verify_signature(self):
        """
        验证许可证签名
        :return: 签名有效返回True，否则返回False
        """
        if not self.license_data or 'signature' not in self.license_data:
            return False

        # 获取签名
        original_signature = self.license_data['signature']

        # 创建数据的副本用于验证
        data_copy = self.license_data.copy()
        data_copy.pop('signature', None)

        # 准备数据字符串
        data_str = json.dumps(data_copy, sort_keys=True)

        # 获取密钥
        secret_key = self._load_key()

        # 计算签名
        calculated_signature = hmac.new(
            secret_key.encode(),
            data_str.encode(),
            hashlib.sha256
        ).hexdigest()

        # 比较签名
        return hmac.compare_digest(original_signature, calculated_signature)

    def _verify_hardware(self):
        """
        验证当前硬件是否匹配许可证
        :return: 匹配返回True，否则返回False
        """
        if not self.license_data or 'hardware_id' not in self.license_data:
            return False

        # 获取许可证中的硬件ID
        licensed_hardware_id = self.license_data['hardware_id']

        # 获取当前硬件ID
        fingerprint = HardwareFingerprint()
        fingerprint.collect_all_data()
        current_hardware_id = fingerprint.get_fingerprint_id()

        # 比较硬件ID
        return licensed_hardware_id == current_hardware_id

    def _verify_expiry(self):
        """
        验证许可证是否过期
        :return: 未过期返回True，已过期返回False
        """
        if not self.license_data or 'expiry_date' not in self.license_data:
            return False

        try:
            # 解析过期日期
            expiry_date = datetime.datetime.strptime(
                self.license_data['expiry_date'],
                '%Y-%m-%d %H:%M:%S'
            )

            # 比较当前日期和过期日期
            return datetime.datetime.now() <= expiry_date

        except Exception as e:
            print(f"验证过期日期时发生错误: {e}")
            return False

    def validate(self):
        """
        执行完整的许可证验证
        :return: (是否有效, 错误消息)
        """
        # 加载许可证
        if not self.load_license():
            return False, "加载许可证失败"

        # 验证签名
        if not self._verify_signature():
            return False, "许可证签名无效"

        # 验证硬件
        if not self._verify_hardware():
            return False, "软件未授权在此计算机上运行"

        # 验证过期日期
        if not self._verify_expiry():
            return False, "许可证已过期"

        # 所有检查都通过
        return True, "许可证有效"


    def get_license_info(self):
        """
        获取许可证信息
        :return: 许可证信息字典
        """
        if not self.license_data:
            return {}

        # 创建一个副本，排除签名
        info = self.license_data.copy()
        info.pop('signature', None)

        return info



    def _anti_debug(self):
        """反调试检测"""
        if hasattr(sys, 'gettrace') and sys.gettrace():
            logging.critical("检测到调试器!")
            sys.exit(105)


# 用于测试的简单示例
if __name__ == "__main__":
    validator = LicenseValidator(
        license_path="license.dat",
        key_path="license_secret.key"  # 生产环境中不应使用外部密钥文件
    )

    valid, message = validator.validate()

    if valid:
        print("许可证验证成功!")

        # 显示许可证信息
        license_info = validator.get_license_info()
        print("\n许可证信息:")
        for key, value in license_info.items():
            print(f"{key}: {value}")
    else:
        print(f"许可证验证失败: {message}")


#执行程序验证license