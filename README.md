# 简介

把mindmaster导出的excel大纲转换为类似“幕布”的大纲格式，以便复制到印象笔记保存。

# 背景

把mindmaster画的思维导图备份到印象笔记中，但是他自带的导出不太令人满意。个人比较喜欢“幕布”的大纲，所以这个项目实现mindmaster导出类似“幕布”的大纲。

# 步骤

1. 准备工作：把项目先用maven打包（需要maven环境），得到 mindMaster-to-note.jar
2. 在mindmaster中先把大纲导出为 Excel 文件
3. 把导出的 Excel 文件和 mindMaster-to-note.jar 放在一个目录
4. 命令窗口执行（需要有Java运行环境）  java  -jar  mindMaster-to-note.jar  \*\*\*.xlsx  （\*\*\*.xlsx为导出的 Excel 文件）
5. 这时当前目录会生成一个html文件
6. 用浏览器打开上面生成的HTML文件，全选复制到印象笔记中即可。

# 其他

把mindmaster导图备份到印象笔记可以模仿xmind，包括三个部分：

* 源文件(*.emmx)*

* 思维导图图片（我比较喜欢svg格式的）
* 大纲

源文件和图片直接用mindMaster导出，拖拽进印象笔记即可。