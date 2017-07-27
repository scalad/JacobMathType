# JacobMathType

JACOB是一个 Java到微软的COM接口的桥梁。使用JACOB允许任何JVM访问COM对象，从而使JAVA应用程序能够调用COM对象。如果你要对 MS Word、Excel 进行处理，JACOB 是一个好的选择。JACOB目前已经成为sourceforge[https://sourceforge.net/projects/jacob-project/?source=directory](https://sourceforge.net/projects/jacob-project/?source=directory)的一个开源项目，官网地址为[http://danadler.com/jacob/](http://danadler.com/jacob/),现在最新的版本是jacob-1.18,你可以在[https://sourceforge.net/projects/jacob-project/](https://sourceforge.net/projects/jacob-project/)上找到最新的jacob.jar包和jacob.dll库.

把jacob下载下来以后解压，里面有两个文件一个是dll另一个是jar文件。将dll文件放入C:/WINDOWS/system32下面，然后把jar文件加入你要用的工程里面就可以使用jacob了.

要使用jacob重要的是要理解VBA的用法，因为jacob其实就是VBA的一个java接口，它提供了一种方法让你可以调用VBA。所以在你要是VBA以前最好先去MSDN上面查看一下office 的reference 上面有一个文档如何创建，打开，保存关闭等功能。我在学习jacob用法的时候就是因为不懂VBA，在哪里胡乱的试，浪费了不少时间。最后还是在msdn上才找到了我要的东西。所以你要用jacob一定要先了解VBA。

