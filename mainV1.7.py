import sys
import logging
from PyQt5.QtWidgets import QApplication, QMessageBox
from license_validator import LicenseValidator

# 延迟导入主程序模块
from STARCCM_Simulation_automation_V7_1 import SimulationConfigWindow


def setup_logging():
    """配置日志记录"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("startup.log"),
            logging.StreamHandler()
        ]
    )


def main():
    setup_logging()
    app = QApplication(sys.argv)

    try:
        # 显示启动加载界面
        splash = QMessageBox()
        splash.setWindowTitle("系统初始化")
        splash.setText("<b>正在验证软件许可证...</b>")
        splash.setStyleSheet("""
            QLabel { 
                font-size: 18px; 
                min-width: 400px; 
                min-height: 150px;
            }
        """)
        splash.show()
        QApplication.processEvents()

        # 初始化验证器
        validator = LicenseValidator(
            license_path="license.dat",
            key_path="license_secret.key"  # 使用硬编码密钥
        )

        # 执行验证
        valid, message = validator.validate()

        if valid:
            splash.close()
            logging.info("许可证验证通过，启动主程序")

            # # 延迟导入主程序模块
            # from STARCCM_Simulation_automation_V6_0 import SimulationConfigWindow

            # 启动主界面
            window = SimulationConfigWindow(validator)
            window.show()

            # 启动心跳检查（每小时检查一次）
            # validator.start_heartbeat(interval=20)

            sys.exit(app.exec_())
        else:
            error_code = hash(message) & 0xFFFF
            QMessageBox.critical(
                None,
                "许可证验证失败",
                f"""<b>错误信息：</b>{message}
                <br><b>错误代码：</b>0X{error_code:04X}
                <br><br>请联系技术支持获取帮助""",
                buttons=QMessageBox.Ok
            )
            sys.exit(101)

    except Exception as e:
        logging.error(f"启动过程中发生致命错误: {str(e)}", exc_info=True)
        QMessageBox.critical(
            None,
            "系统错误",
            f"""<b>程序初始化失败：</b>{str(e)}
            <br><br>请检查以下内容：
            <br>1. 确保所有依赖文件完整
            <br>2. 确认许可证文件存在
            <br>3. 检查系统日期时间是否正确""",
            buttons=QMessageBox.Ok
        )
        sys.exit(102)


if __name__ == "__main__":
    main()
