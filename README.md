## 代码功能
用于快速计算bug表单（类似于从PMS导出的bug表单，需要表单中至少含有严重程度和bug状态两项）中bug的DI值。
## 使用方法
此代码可以通过右键点击excel文件，然后计算其中的bug的DI值。
为了能实现右键点击的功能，需要将run_calculate_di.desktop存放到：/usr/share/deepin/dde-file-manager/oem-menuextensions（UOS操作系统中）

运行前需要对CulculateDI.py和run_calculate_di.desktop添加执行权限:
chmod +x filepath