首先，感谢您使用妖王牌qq消息轰炸器！
#注册消息：
公司：港田路360号西马神奇动物园
程序员：MorganFish
邮箱：MorganFish0508@163.com
GitHub：https://www.github.com/MorganNotFound
以下便是本产品的使用方式：
step1：
解压文件压缩包（不过您既然看到了此文件说明您已经完成了此步骤）
step2：
右键文件中的boom.VBS然后点击“编辑”
#打开后是以下代码：
   Set WshShell=WScript.CreateObject("WScript.Shell")
   WshShell.AppActivate"嘿嘿"            '这个"嘿嘿"不用更改
   for i=1 to 20                         '这一句中的20是发送的消息数量
   WScript.Sleep 10                      '这个10是相邻两条消息间隔的毫秒数
   WshShell.SendKeys"^v"
   WshShell.SendKeys"%s"
   Next
保存后关闭，然后右键文件打开“属性”，复制文件位置，如C:\Users\Administrator\Desktop\QQ消息轰炸终版\boom.VBS    （一般来说复制之后不会携带\boom.VBS，可以之后再加上）
step3：
打开轰炸器.exe
#确保您的qq处于打开状态
注意：一定要看！！！这一段的使用方法很重要，与实际情况并不相同，请注意！！！
   （一）输入刚才的boom.VBS地址；
   （二）输入您想要轰炸的qq号；
   （三）将您想要轰炸发送的内容输入，然后全选--复制
   （四）不用管那个发送数量与间隔，本人手残做了两个废物模块（水平有限，勿喷
            #不过如果您有兴趣可以考虑改进程序，实现输入即可达成目的
   （五）点击开始即可

但是，有没有发现如果第二次打开程序boom.VBS的地址位置又被重置了？
怎么办呢？
很简单，只要您在工程里将Text5.Text里的内容改为您的boom.VBS地址然后重新导出工程.exe即可
再次感谢您的使用！
（我还是个学生，手残勿喷）