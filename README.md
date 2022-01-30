使用方法：

1.导出文件夹内所有excel文件

  Excel2Code.exe -dir Excels
  
2.导出指定excel文件

  Excel2Code.exe Excels/test.xlsx Excels/test2.xlsx
  
------------------
·以#开头的Sheet，会被识别为基础类型集合，第一列为类型，第二列为变量名，第三列为内容，内容为 int/string/bool 时，第一列可以为空，会自己推测类型
------------------
·Sheet名为一个单词时，这个Sheet会被识别为自定义类，第一列为类型，第二列为变量名
------------------
·Sheet名为“单词空格单词”格式时，这个Sheet会被识别为数据，Sheet名为“类名 变量名”，第一行被识别为类里的各个变量，当只有第二行一个数据时，这个Sheet会被识别为成员变量，当有多行数据时，这个Sheet会被识别为这个类的列表，即List<类>。
------------------
Excel2CSharp/Excel2Code/bin/Debug/net5.0/Excels/目录下有目前支持的所有写法的例子，生成的C#文件在Excel2CSharp/Excel2Code/bin/Debug/net5.0/Codes目录下
