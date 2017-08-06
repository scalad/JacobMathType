# JacobMathType

### Jacob ###
JACOB是一个 Java到微软的COM接口的桥梁。使用JACOB允许任何JVM访问COM对象，从而使JAVA应用程序能够调用COM对象。如果你要对 MS Word、Excel 进行处理，JACOB 是一个好的选择。JACOB目前已经成为sourceforge[https://sourceforge.net/projects/jacob-project/?source=directory](https://sourceforge.net/projects/jacob-project/?source=directory)的一个开源项目，官网地址为[http://danadler.com/jacob/](http://danadler.com/jacob/),现在最新的版本是jacob-1.18,你可以在[https://sourceforge.net/projects/jacob-project/](https://sourceforge.net/projects/jacob-project/)上找到最新的jacob.jar包和jacob.dll库,使用的时候还需要注意这两个东西的版本需要一致，而且还分32位和64位的，它的位数和JDK的位数有关，和操作系统的位数无关。

### 早期绑定和晚期绑定 ###
调用 COM 中暴露出来的方法，主要有两种机制：早期绑定和晚期绑定。

早期绑定显式的声明对象、数据类型等，编译器获取了足够的信息在编译期进行链接和优化，这样通常可以获得更好的性能，通过这种机制实现 Bridge 调用可以参考 IBM 的 RJCB 项目，它提供了一套高性能的解决方案。当然您需要了解更多 COM 组件的细节，虽然框架为您完成了大部分的生成 Bridge 代码的工作，但是总体来说，编码工作量还是偏大，编码难度也比较高，而且 RJCB 仅支持那些提供早期绑定的 vtable 接口的 COM API。

而晚期绑定方式是通过 IDispatch接口来实现，类似 Java 的反射机制，您可以按照名称或者 ID 进行方法调用，这种设计主要目的是支持脚本语言操作 COM，因为脚本是解释执行的，通常都不支持指针也就没有 C++ 中的 vtable 机制，无法实现早期绑定。这种方式的缺点是在一定程度上性能较差，由于是在运行时按照名字或者 ID 进行对象的调用，只有运行时才确切知道调用的是哪个对象，哪个方法，这样必然带来一定的开销和延迟。但是这种方式的优点也是非常明显的，简单、灵活，您可以不必关注动态链接库的细节，可以非常快地完成代码开发工作。

### 核心 ###
JACOB 开源项目提供的是一个 JVM 独立的自动化服务器实现，其核心是基于 JNI 技术实现的 Variant, Dispatch 等接口，设计参考了 Microsoft VJ++ 内置的通用自动化服务器，但是 Microsoft 的实现仅仅支持自身的 JVM。通过 JACOB，您可以方便地在 Java 语言中进行晚期绑定方式的调用。

下图是一个对 JACOB 结构的简单说明

![](https://github.com/scalad/JacobMathType/blob/master/doc/image/image001.jpg)

把jacob下载下来以后解压，里面有两个文件一个是dll另一个是jar文件。将dll文件放入C:/WINDOWS/system32下面或者是jdk所在的路径下，然后把jar文件加入你要用的工程里面就可以使用jacob了.

![](https://github.com/scalad/JacobMathType/blob/master/doc/image/jacob_path.png)

要使用jacob重要的是要理解VBA的用法，因为jacob其实就是VBA的一个java接口，它提供了一种方法让你可以调用VBA。所以在你要是VBA以前最好先去MSDN上面查看一下office 的reference 上面有一个文档如何创建，打开，保存关闭等功能。我在学习jacob用法的时候就是因为不懂VBA，在哪里胡乱的试，浪费了不少时间。最后还是在msdn上才找到了我要的东西。所以你要用jacob一定要先了解VBA。

jacob用来调用实现COM接口的dll。根据分析jacob提供的类，发现com.jacob.com.Dispatch有invoke方法。可以猜到使用java反射机制的方式调用dll。那么你只要dll的api，以传入类名、方法名、参数方式就可以调用dll。

### 常用类以及方法 ###
* **ComThread**：com组件管理，用来初始化com线程，释放线程，所以会在操作office之前使用，操作完成再使用。 
* **ActiveXComponent**：创建office的一个应用，比如你操作的是word还是excel 
* **Dispatch**：调度处理类，封装了一些操作来操作office，里面所有的可操作对象基本都是这种类型，所以jacob是一种链式操作模式，就像StringBuilder对象，调用append()方法之后返回的还是StringBuilder对象 
* **Variant**：封装参数数据类型，因为操作office是的一些方法参数，可能是字符串类型，可能是数字类型，虽然都是1，但是不能通过，可以通过Variant来进行转换通用的参数类型，new Variant(1),new Variant("1") 
* **Dispatch的几种静态方法**：这些方法就是要用来操作office的。
	* call( )方法：调用COM对象的方法，返回Variant类型值。
	* invoke( )方法：和call方法作用相同，但是不返回值。 
	* get( )方法：获取COM对象属性，返回variant类型值。 
	* put( )方法：设置COM对象属性。

以上方法中有的有很多重载方法，调用不同的方法时需要放置不同的参数，至于哪些参数代表什么意思，具体放什么值，就需要参考vba代码了，仅靠jacob是无法进行变成的。 

**Variant对象的toDispatch()方法**：将以上方法返回的Variant类型转换为Dispatch，进行下一次链式操作 

### MathType数学公式编辑器 ###
MathType 是由美国Design Science公司开发的功能强大的数学公式编辑器，它同时支持Windows和Macintosh 操作系统，与常见的文字处理软件和演示程序配合使用，能够在各种文档中加入复杂的数学公式和符号.

MathType与Office文档完美结合，显示效果超好，比Office自带的公式编辑器要强大很多,关于MathType的安装这里不再介绍，安装完之后MathType犹如一个插件嵌入在Word中，如下图：

![](https://github.com/scalad/JacobMathType/blob/master/doc/image/mathType_word.png)

接下来我们演示如何使用Jacob如何调用MathType的API，开始之前，请确保Window服务中`DCom Server Process Launcher`这个服务已经开启，不然等下会报错`com.jacob.com.ComFailException: Can't co-create object`.

### 具体操作 ###
#### 1、初始化com线程 ####

使用jacob之前使用下面的语句来初始化com线程，大概意思是打开冰箱门，准备放大象
```Java
ComThread.InitSTA(); 
```

使用完成后使用下面的语句来关闭现场,大概意思是关上冰箱门

```
ComThread.Release();
```
#### 2、创建应用程序对象，设置参数，得到文档集合 ####
操作一个文档之前，我们必须要创建一个应用对应，比如是mathType，或者是word还是excel，设置一些文档应用的参数，得到文档集合对象，（大家应该知道word是Documents，excel是WorkBooks）

```Java
//word
ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
//excel
ActiveXComponent wordApp = new ActiveXComponent("Excel.Application");
//设置应用操作是文档不在明面上显示，只在后台静默处理。  
wordApp.setProperty("Visible", new Variant(false));
```