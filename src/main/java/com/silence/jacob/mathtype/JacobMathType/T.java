package com.silence.jacob.mathtype.JacobMathType;

public class T {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		MSWordManager ms = new MSWordManager(false);
		// 生成一个MSwordManager对象,并且设置显示Word程序
		ms.createNewDocument();
		// 创建一个新的.doc文件
		ms.insertText("Test jacob");
		// 插入文本
		ms.save("c:\\1.doc");
		// 保存.doc文件
		ms.close();
		ms.closeDocument();
	}

}