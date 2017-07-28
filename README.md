# JacobMathType

JACOB是一个 Java到微软的COM接口的桥梁。使用JACOB允许任何JVM访问COM对象，从而使JAVA应用程序能够调用COM对象。如果你要对 MS Word、Excel 进行处理，JACOB 是一个好的选择。JACOB目前已经成为sourceforge[https://sourceforge.net/projects/jacob-project/?source=directory](https://sourceforge.net/projects/jacob-project/?source=directory)的一个开源项目，官网地址为[http://danadler.com/jacob/](http://danadler.com/jacob/),现在最新的版本是jacob-1.18,你可以在[https://sourceforge.net/projects/jacob-project/](https://sourceforge.net/projects/jacob-project/)上找到最新的jacob.jar包和jacob.dll库.

有关自动化的更多细节，您可以参考相关文档和书籍，我们仅做简单介绍。调用 COM 中暴露出来的方法，主要有两种机制：早期绑定和晚期绑定。

早期绑定显式的声明对象、数据类型等，编译器获取了足够的信息在编译期进行链接和优化，这样通常可以获得更好的性能，通过这种机制实现 Bridge 调用可以参考 IBM 的 RJCB 项目，它提供了一套高性能的解决方案。当然您需要了解更多 COM 组件的细节，虽然框架为您完成了大部分的生成 Bridge 代码的工作，但是总体来说，编码工作量还是偏大，编码难度也比较高，而且 RJCB 仅支持那些提供早期绑定的 vtable 接口的 COM API。

而晚期绑定方式是通过 IDispatch接口来实现，类似 Java 的反射机制，您可以按照名称或者 ID 进行方法调用，这种设计主要目的是支持脚本语言操作 COM，因为脚本是解释执行的，通常都不支持指针也就没有 C++ 中的 vtable 机制，无法实现早期绑定。这种方式的缺点是在一定程度上性能较差，由于是在运行时按照名字或者 ID 进行对象的调用，只有运行时才确切知道调用的是哪个对象，哪个方法，这样必然带来一定的开销和延迟。但是这种方式的优点也是非常明显的，简单、灵活，您可以不必关注动态链接库的细节，可以非常快地完成代码开发工作。

JACOB 开源项目提供的是一个 JVM 独立的自动化服务器实现，其核心是基于 JNI 技术实现的 Variant, Dispatch 等接口，设计参考了 Microsoft VJ++ 内置的通用自动化服务器，但是 Microsoft 的实现仅仅支持自身的 JVM。通过 JACOB，您可以方便地在 Java 语言中进行晚期绑定方式的调用。

下图是一个对 JACOB 结构的简单说明

![](https://github.com/scalad/JacobMathType/blob/master/doc/image/image001.jpg)

把jacob下载下来以后解压，里面有两个文件一个是dll另一个是jar文件。将dll文件放入C:/WINDOWS/system32下面，然后把jar文件加入你要用的工程里面就可以使用jacob了.

要使用jacob重要的是要理解VBA的用法，因为jacob其实就是VBA的一个java接口，它提供了一种方法让你可以调用VBA。所以在你要是VBA以前最好先去MSDN上面查看一下office 的reference 上面有一个文档如何创建，打开，保存关闭等功能。我在学习jacob用法的时候就是因为不懂VBA，在哪里胡乱的试，浪费了不少时间。最后还是在msdn上才找到了我要的东西。所以你要用jacob一定要先了解VBA。

