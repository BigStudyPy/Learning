#环境配置：
        python==3.8
        selenium==4.1.0
        xlrd==2.0.1
		Pillow==8.4.0
        numpy==1.21.4
		
        chrome==96.0.4664.45
        chromedriver对应版本
    现有Python一键安装
        pip install selenium==4.1.0
        pip install xlrd==2.0.1
        pip install numpy==1.21.4
    conda环境配置
        conda create -n bigstudy python=3.8
        conda activate bigstudy
        conda install selenium==4.1.0
        conda install xlrd==2.0.1
        conda install pillow==8.4.0
        conda install numpy==1.21.4
    chrome浏览器官网下载：https://www.google.cn/chrome/
    chromedriver官方下载：http://chromedriver.storage.googleapis.com/index.html
    chromedriver镜像下载：https://npm.taobao.org/mirrors/chromedriver/
    
    最后，把chromedriver解压到chrome浏览器所在的目录，将路径加入path环境变量
#使用方法
    执行big_study.py文件即可
	此程序会读取表格中数据进行青年大学习的自动学习与截图,并最终以邮件的方式发送到团支书邮箱中
	截图中的字体是在网上随便下载的一个免费商用字体，想换的朋友可以自行替换
	注意：大学习密码不能是纯数字
	大学习账号一般是手机号，密码可在共青团微信公众号中：大学习->注册用户登录->我的->用户信息修改
	中进行修改
	记得每一期大学习自己手动换网址，具体方法是更改chrome.get（''）中的网址

    
        
        