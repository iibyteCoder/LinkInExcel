# LinkInExcel
标记excel中失效的链接

配置：
1. 配置文件为config/configuration.json 
2. 若使用代理，则需要在的PROXIES加上本地对应端口的ip地址
3. 修改THREAD_COUNT，调整线程数

食用方法：
1. 将需要检测的excel文件放入inputFile文件夹
2. 在本项目文件夹根目录呼出cmd窗口
3. 在窗口命令行输入
    .\linkChecker.exe '需要标记的文件名XXX.xlsx'
    按下回车键，文件开始执行
4. 等待执行结束，在outputFile文件夹内会生成带有标记的文件
    文件名格式：原文件名_XXXX_XX_XX_XX_XX_XX.xlsx (年月日时分秒)
