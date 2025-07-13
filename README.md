此应用可以批量快速的更新客人的PIN FISCONLINE密码

## Author
-Qin_5352 - https://github.com/VEQ25

https://googlechromelabs.github.io/chrome-for-testing/#stable
下载相同版本的Chrome和ChromeDriver,把chrome-win64和ChromeDriver.exe放在Reset_Password_Pin_Fisconline的目录下

以下是使用方法:

1) 在Reset_Password.xlsx里输入想要更新密码客人的CODICE FISCALE - CODICE PIN - PASSWORD INIZIALE- NUOVA PASSWORD, 
    并且把Formato celle全都改成Testo(这样CODICE PIN开头是0 也会正常显示0，CODICE PIN是10位数的，运行之前可以先检查一下), SALVA EXCEL。

2) 双击Reset_Password.exe 运行 (如果显示更新失败，检查EXCEL - Reset_Password.xlsx里的数据是否填写错误，然后删掉已经成功更新的客人，重新运行，或者手动更改密码)，
    运行期间不要打开EXCEL，最后按Entrer结束运行。

3) 运行结束后会生成日志保存在logs文件夹里查看更新记录，SUCCESS表示成功，FAIL表示失败
