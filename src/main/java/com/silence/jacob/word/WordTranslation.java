package com.silence.jacob.word;

import java.io.IOException;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sience.jacob.util.Utils;

public class WordTranslation {

	public void translate() {
		//初始化COM线程
		ComThread.InitSTA();
		
		//创建应用程序对象，设置参数，得到文档集合 ,得到文档集合对象
		ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
		//设置应用操作是文档不在明面上显示，只在后台静默处理。
		wordApp.setProperty("Visible", new Variant(false));
		
		//打开文档集合
		Dispatch document = wordApp.getProperty("Documents").toDispatch();
		
		//打开文档
		Dispatch doc = Dispatch.call(document, "Open").getDispatch();
		
		//退出Word Application
		wordApp.invoke("Quit", new Variant() {});
		//释放com线程
		ComThread.Release();
	}

	public static void main(String[] args) throws IOException {
		Utils.loadLibrary(System.getProperty("user.dir"));
		//初始化COM线程
		ComThread.InitSTA();
		ActiveXComponent mathTypeApp = new ActiveXComponent("Equation.DSMT4");
		Dispatch  dispatch = mathTypeApp.getObject();
		
		Variant v = Dispatch.call(dispatch, "DefaultIcon");
	    System.out.println(v);
		//释放com线程
		ComThread.Release();
	}
}
