# 工程造价对比分析工具
  <br>快速安装：<br>
  <br>1、确保已安装python环境，下载code文件夹中所有内容<br>
  <br>2、运行run.bat安装所需扩展库<br>
  <br>3、运行cost_compare.py，开始分析<br>

<br>简明操作指引<br>
<br>![操作界面](https://github.com/jason449154834/cost-comparison-tools/blob/main/pic/%E6%93%8D%E4%BD%9C%E7%95%8C%E9%9D%A2.png)<br>
<br>基准项目、对比项目：选择2个对比项目的xlsx格式文件<br>
<br>序号、项目编码、项目名称、项目特征、计量单位、工程量、单价：在文件中，首次出现在第几列。例：项目特征在D及E列（被合并单元格），则设置为3（A为0，B为1，C为2，依此类推）<br>
<br>最相近清单数：在对比项目中，选择出与基准项目最相近清单的数量<br>
<br>相近概率低于%则不对比：选择出与基准项目最相近清单均低于此值，则对比工程量设置为0<br>
<br>取值方法：选择出与基准项目最相近清单后，按最相近值或者平均值进行对比<br>
<br>设置好后就开始分析吧！<br>


<br>Q&A<br>
<br>Q：为什么不打包成EXE文件？ A：python的pyinstaller有太多莫名其妙的bug，故只发布源代码<br>
