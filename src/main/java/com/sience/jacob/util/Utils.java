package com.sience.jacob.util;

import java.io.IOException;
import java.nio.file.Paths;

import com.jacob.com.LibraryLoader;

public class Utils {

	/**
	 * 加载DLL文件并设置到系统环境中
	 * @param appDir
	 * @throws IOException
	 */
	public static void loadLibrary(final String appDir) throws IOException {
		final String libFile = "amd64".equals(System.getProperty("os.arch")) 
				? "/jacob-1.18-x64.dll" : "/jacob-1.18-x86.dll";
		System.setProperty(LibraryLoader.JACOB_DLL_PATH, Paths.get(appDir, "lib", libFile).toString());
		LibraryLoader.loadJacobLibrary();
	}
}
