import hashlib
import platform
import subprocess
import re
import os
import json



class HardwareFingerprint:
    def __init__(self):
        self.fingerprint_data = {}
        self.unique_id = None

    def collect_all_data(self):
        """收集所有可用的硬件指纹数据"""
        self.get_system_info()
        self.get_cpu_info()
        self.get_bios_info()
        self.get_motherboard_info()
        self.get_disk_info()
        self.get_mac_address()

        # 生成唯一指纹
        self.generate_unique_id()

        return self.fingerprint_data

    def get_system_info(self):
        """获取基本系统信息"""
        self.fingerprint_data['system'] = {
            'platform': platform.system(),
            'platform_release': platform.release(),
            'platform_version': platform.version(),
            'architecture': platform.machine(),
            'hostname': platform.node(),
            'processor': platform.processor(),
        }

    def get_cpu_info(self):
        """获取CPU信息"""
        cpu_info = {}

        if platform.system() == "Windows":
            # Windows系统下获取CPU信息
            try:
                output = subprocess.check_output(
                    "wmic cpu get ProcessorId", shell=True).decode().strip()
                processor_id = re.search(r'ProcessorId\s*(.*)', output).group(1).strip()
                cpu_info['processor_id'] = processor_id
            except:
                cpu_info['processor_id'] = "Unknown"

            try:
                output = subprocess.check_output(
                    "wmic cpu get Name,NumberOfCores,MaxClockSpeed", shell=True).decode()
                cpu_name = re.search(r'Name\s*(.*)', output).group(1).strip()
                cores = re.search(r'NumberOfCores\s*(.*)', output).group(1).strip()
                cpu_info['name'] = cpu_name
                cpu_info['cores'] = cores
            except:
                pass

        elif platform.system() == "Linux":
            # Linux系统下获取CPU信息
            try:
                with open('/proc/cpuinfo', 'r') as f:
                    cpuinfo = f.read()

                # 尝试找到CPU ID，物理ID或序列号
                matches = re.search(r'physical id\s+:\s+(\d+)', cpuinfo)
                if matches:
                    cpu_info['physical_id'] = matches.group(1)

                matches = re.search(r'model name\s+:\s+(.*)', cpuinfo)
                if matches:
                    cpu_info['model_name'] = matches.group(1)

                matches = re.search(r'cpu cores\s+:\s+(\d+)', cpuinfo)
                if matches:
                    cpu_info['cores'] = matches.group(1)
            except:
                pass

        elif platform.system() == "Darwin":  # macOS
            try:
                output = subprocess.check_output(
                    "sysctl -n machdep.cpu.brand_string", shell=True).decode().strip()
                cpu_info['brand'] = output

                output = subprocess.check_output(
                    "sysctl -n hw.physicalcpu", shell=True).decode().strip()
                cpu_info['physical_cores'] = output
            except:
                pass

        self.fingerprint_data['cpu'] = cpu_info

    def get_bios_info(self):
        """获取BIOS信息"""
        bios_info = {}

        if platform.system() == "Windows":
            try:
                output = subprocess.check_output(
                    "wmic bios get Manufacturer,SerialNumber,Version,ReleaseDate",
                    shell=True).decode()

                manufacturer = re.search(r'Manufacturer\s*(.*)', output)
                if manufacturer:
                    bios_info['manufacturer'] = manufacturer.group(1).strip()

                serial = re.search(r'SerialNumber\s*(.*)', output)
                if serial:
                    bios_info['serial_number'] = serial.group(1).strip()

                version = re.search(r'Version\s*(.*)', output)
                if version:
                    bios_info['version'] = version.group(1).strip()

                date = re.search(r'ReleaseDate\s*(.*)', output)
                if date:
                    bios_info['release_date'] = date.group(1).strip()
            except:
                pass

        elif platform.system() == "Linux":
            try:
                if os.path.exists('/sys/class/dmi/id/bios_vendor'):
                    with open('/sys/class/dmi/id/bios_vendor', 'r') as f:
                        bios_info['vendor'] = f.read().strip()

                if os.path.exists('/sys/class/dmi/id/bios_version'):
                    with open('/sys/class/dmi/id/bios_version', 'r') as f:
                        bios_info['version'] = f.read().strip()

                if os.path.exists('/sys/class/dmi/id/bios_date'):
                    with open('/sys/class/dmi/id/bios_date', 'r') as f:
                        bios_info['date'] = f.read().strip()
            except:
                pass

        self.fingerprint_data['bios'] = bios_info

    def get_motherboard_info(self):
        """获取主板信息"""
        motherboard_info = {}

        if platform.system() == "Windows":
            try:
                output = subprocess.check_output(
                    "wmic baseboard get Manufacturer,Product,SerialNumber",
                    shell=True).decode()

                manufacturer = re.search(r'Manufacturer\s*(.*)', output)
                if manufacturer:
                    motherboard_info['manufacturer'] = manufacturer.group(1).strip()

                product = re.search(r'Product\s*(.*)', output)
                if product:
                    motherboard_info['product'] = product.group(1).strip()

                serial = re.search(r'SerialNumber\s*(.*)', output)
                if serial:
                    motherboard_info['serial_number'] = serial.group(1).strip()
            except:
                pass

        elif platform.system() == "Linux":
            try:
                if os.path.exists('/sys/class/dmi/id/board_vendor'):
                    with open('/sys/class/dmi/id/board_vendor', 'r') as f:
                        motherboard_info['vendor'] = f.read().strip()

                if os.path.exists('/sys/class/dmi/id/board_name'):
                    with open('/sys/class/dmi/id/board_name', 'r') as f:
                        motherboard_info['name'] = f.read().strip()

                if os.path.exists('/sys/class/dmi/id/board_serial'):
                    with open('/sys/class/dmi/id/board_serial', 'r') as f:
                        motherboard_info['serial'] = f.read().strip()
            except:
                pass

        self.fingerprint_data['motherboard'] = motherboard_info

    def get_disk_info(self):
        """获取硬盘信息"""
        disk_info = {}

        if platform.system() == "Windows":
            try:
                output = subprocess.check_output(
                    "wmic diskdrive get Model,SerialNumber,Size",
                    shell=True).decode()

                lines = output.strip().split('\n')
                if len(lines) > 1:  # 第一行是标题
                    # 解析第一个硬盘的信息
                    disk_line = ' '.join(lines[1].split())  # 规范化空格
                    parts = disk_line.split(' ')

                    # 假设最后一个部分是Size，倒数第二个是SerialNumber
                    if len(parts) >= 3:
                        size_index = len(parts) - 1
                        serial_index = len(parts) - 2
                        model = ' '.join(parts[:serial_index])

                        disk_info['model'] = model
                        disk_info['serial_number'] = parts[serial_index]
                        disk_info['size'] = parts[size_index]
            except:
                pass

        elif platform.system() == "Linux":
            try:
                # 尝试使用lsblk获取磁盘信息
                output = subprocess.check_output(
                    "lsblk -d -o NAME,SERIAL,SIZE --nodeps",
                    shell=True).decode()

                lines = output.strip().split('\n')
                if len(lines) > 1:
                    for line in lines[1:]:  # 跳过标题行
                        parts = line.split()
                        if len(parts) >= 3:
                            name = parts[0]
                            serial = parts[1]
                            size = parts[2]

                            disk_info[name] = {
                                'serial': serial,
                                'size': size
                            }
                            break  # 只获取第一个磁盘
            except:
                # 尝试使用hdparm
                try:
                    # 找到第一个硬盘设备
                    output = subprocess.check_output(
                        "ls /dev/sd* | grep -E '^/dev/sd[a-z]$' | head -1",
                        shell=True).decode().strip()

                    if output:
                        # 获取序列号
                        serial_output = subprocess.check_output(
                            f"hdparm -I {output} | grep 'Serial Number'",
                            shell=True).decode().strip()

                        if "Serial Number" in serial_output:
                            serial = serial_output.split(':')[1].strip()
                            disk_info['serial_number'] = serial
                except:
                    pass

        self.fingerprint_data['disk'] = disk_info

    def get_mac_address(self):
        """获取网卡MAC地址"""
        mac_addresses = []

        # 获取所有网卡的MAC地址
        for interface in self._get_interfaces():
            try:
                mac = self._get_interface_mac(interface)
                if mac and mac != "00:00:00:00:00:00":
                    mac_addresses.append((interface, mac))
            except:
                pass

        self.fingerprint_data['mac_addresses'] = mac_addresses

    def _get_interfaces(self):
        """获取网络接口列表"""
        interfaces = []

        if platform.system() == "Windows":
            try:
                output = subprocess.check_output(
                    "ipconfig /all", shell=True).decode()
                interfaces = re.findall(r'adapter (.+):', output)
            except:
                pass
        elif platform.system() in ["Linux", "Darwin"]:
            try:
                if platform.system() == "Linux":
                    output = subprocess.check_output("ls /sys/class/net/", shell=True).decode()
                else:  # macOS
                    output = subprocess.check_output(
                        "networksetup -listallhardwareports | grep Device | awk '{print $2}'", shell=True).decode()
                interfaces = output.strip().split()
            except:
                pass

        return interfaces

    def _get_interface_mac(self, interface):
        """获取特定网络接口的MAC地址"""
        if platform.system() == "Windows":
            try:
                output = subprocess.check_output(
                    f"ipconfig /all", shell=True).decode()
                section_pattern = re.escape(
                    interface) + r':[.\s\S]*?Physical Address[.\s\S]*?([\da-fA-F]{2}(-|:)[\da-fA-F]{2}(-|:)[\da-fA-F]{2}(-|:)[\da-fA-F]{2}(-|:)[\da-fA-F]{2}(-|:)[\da-fA-F]{2})'
                match = re.search(section_pattern, output)
                if match:
                    return match.group(1)
            except:
                pass
        elif platform.system() == "Linux":
            try:
                with open(f'/sys/class/net/{interface}/address', 'r') as f:
                    return f.read().strip()
            except:
                pass
        elif platform.system() == "Darwin":  # macOS
            try:
                output = subprocess.check_output(
                    f"ifconfig {interface} | grep ether | awk '{{print $2}}'",
                    shell=True).decode()
                return output.strip()
            except:
                pass

        return None

    def generate_unique_id(self):
        """根据收集的硬件信息生成唯一标识符"""
        fingerprint_str = json.dumps(self.fingerprint_data, sort_keys=True)
        hash_object = hashlib.sha256(fingerprint_str.encode())
        self.unique_id = hash_object.hexdigest()
        self.fingerprint_data['unique_id'] = self.unique_id

    def save_to_file(self, filename='hardware_fingerprint.json'):
        """保存硬件指纹信息到文件"""
        with open(filename, 'w') as f:
            json.dump(self.fingerprint_data, f, indent=4)
        print(f"硬件指纹信息已保存到 {filename}")

    def get_fingerprint_id(self):
        """返回唯一标识符"""
        if not self.unique_id:
            self.generate_unique_id()
        return self.unique_id


if __name__ == "__main__":
    print("收集硬件指纹信息...")
    fingerprint = HardwareFingerprint()
    fingerprint.collect_all_data()

    # 显示唯一ID
    print(f"\n唯一硬件指纹: {fingerprint.get_fingerprint_id()}")

    # 保存到文件
    fingerprint.save_to_file()

    # 显示详细信息
    print("\n收集到的硬件信息摘要:")

    # 系统信息
    sys_info = fingerprint.fingerprint_data.get('system', {})
    print(f"系统: {sys_info.get('platform', 'Unknown')} {sys_info.get('platform_version', '')}")
    print(f"架构: {sys_info.get('architecture', 'Unknown')}")
    print(f"主机名: {sys_info.get('hostname', 'Unknown')}")

    # CPU信息
    cpu_info = fingerprint.fingerprint_data.get('cpu', {})
    cpu_id = cpu_info.get('processor_id') or cpu_info.get('physical_id') or cpu_info.get('brand')
    cpu_name = cpu_info.get('name') or cpu_info.get('model_name')
    print(f"CPU ID: {cpu_id or 'Unknown'}")
    if cpu_name:
        print(f"CPU型号: {cpu_name}")

    # 主板信息
    mb_info = fingerprint.fingerprint_data.get('motherboard', {})
    mb_manufacturer = mb_info.get('manufacturer') or mb_info.get('vendor')
    mb_serial = mb_info.get('serial_number') or mb_info.get('serial')
    print(f"主板制造商: {mb_manufacturer or 'Unknown'}")
    print(f"主板序列号: {mb_serial or 'Unknown'}")

    # BIOS信息
    bios_info = fingerprint.fingerprint_data.get('bios', {})
    bios_version = bios_info.get('version', 'Unknown')
    print(f"BIOS版本: {bios_version}")

    # 硬盘信息
    disk_info = fingerprint.fingerprint_data.get('disk', {})
    if isinstance(disk_info, dict):
        disk_serial = disk_info.get('serial_number')
        if disk_serial:
            print(f"硬盘序列号: {disk_serial}")
        else:
            for disk_name, disk_data in disk_info.items():
                if isinstance(disk_data, dict) and 'serial' in disk_data:
                    print(f"硬盘 {disk_name} 序列号: {disk_data['serial']}")
                    break

    # MAC地址
    mac_addresses = fingerprint.fingerprint_data.get('mac_addresses', [])
    if mac_addresses:
        print(f"首个MAC地址: {mac_addresses[0][1]}")

#执行程序生成指纹
