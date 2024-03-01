# -爬取考试宝题库
1、（可选）可能需要设置系统变量 PYTHONUTF8=1<br />
2、从release下载打包好的程序<br />
3、安装谷歌浏览器并记录下安装位置<br />
4、打开谷歌浏览器访问考试宝并收藏要爬取的题库<br />
5、打开要爬取的题库选择顺序练习并复制地址，回到我的题库页面并且不要关闭浏览器<br />
6、按提示输入需要的选项<br />![image](https://github.com/NICHX/kaoshibao/assets/24547848/93d297aa-a427-4a89-86d8-e79bfe1da43a)<br />
7、点击开始后会开始爬取题库，文件会保存在指定位置的‘paper.txt’并在爬取完成后自动打开文件<br />
8、使用源代码自行打包时可能需要修改gooey/lib/site-packages/gooey/gui/processor.py，否则中文输出异常(参考https://www.cnblogs.com/yunhgu/p/15061756.html)
