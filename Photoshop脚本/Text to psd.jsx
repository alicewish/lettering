/*****************************************************************
 *
 * 此脚本源自
 * 梁进刚
 * 【给汉化者们分享几个小工具】
 * https://tieba.baidu.com/p/3020902369
 * [2014-05-03 20:09]
 *
 *****************************************************************
 *
 * 我的修改：
 * 1、将选项写死在脚本内免得乱弹对话框
 * 2、改进函数
 * 3、增加注释
 * 4、写了一堆配套脚本
 * 5、确认了Photoshop脚本体系bug很多这个事实
 * 6、经IntelliJ IDEA提醒，改用"==="，即全等于，进行判断
 * ————墨问非名
 * [大致从2016年写到2018年]
 *
 *****************************************************************
 *
 * 暗的修改：
 * 1、增加自定义选项
 * [2017年]
 *
 *****************************************************************
 *
 * 备注：
 * 1、建议放置在"C:\Program Files\Adobe\Adobe Photoshop CC 2018\Presets\Scripts"
 * 2、有些变量不能通过函数传递
 * 3、PS脚本不认let关键字定义变量
 * 4、不懂再问
 * ————墨问非名
 *
 *****************************************************************/

//================初始化标尺、字体单位设置================
var originalUnit = preferences.rulerUnits;
preferences.rulerUnits = Units.PIXELS;
var originalTypeUnit = preferences.typeUnits;
preferences.typeUnits = TypeUnits.POINTS;

//================将文本写入数组的函数================
function text2array(read_method) {
    //================输入文本文档================
    if (read_method === '对话框选取') {
        var myTextFile = File.openDialog("打开文件...");
    }
    else if (read_method === '读取特定文档') {
        //在Windows下，路径也必须都用“/”
        //虚拟机:"//psf/Host/Volumes/Mack/-.txt"
        //Win:"//Mac/Dropbox/Test.txt"
        //Mac:"/Users/alicewish/Dropbox/Test.txt"
        var text_path = "//Mac/Dropbox/Test.txt";
        var myTextFile = new File(text_path);
    }
    //================读取文本文档到数组================
    myTextFile.open("r");
    var myLineArray = [];
    while (!myTextFile.eof) {
        myLineArray.push(myTextFile.readln());
    }
    myTextFile.close();
    return myLineArray
}

//================将文本写入图层的函数================
function writeOnLayerS(n, row_len, myLineArray, emptrow) {
    var contents = '';
    if (row_len === 1) {
        contents = myLineArray[emptrow[n] + 1];
    }
    else if (row_len > 1) {
        contents = myLineArray[emptrow[n] + 1];
        for (var j = 1; j < row_len; j++) {
            contents += '\r' + myLineArray[emptrow[n] + j + 1]//'\r'是return，相当于换行
        }
    }
    return contents
}

function main_process(read_method) {
    myLineArray = text2array(read_method);//myLineArray是包含每行文本的数组

    //================读取所有空行所在行数================
    //================尝试写为单独函数但未通过调试================
    var emptrow = [];
    var i = 0;
    for (var lineIndex = 0; lineIndex < myLineArray.length; lineIndex++) {
        if (myLineArray[lineIndex] === "") {
            emptrow[i] = lineIndex;
            i++;
        }
    }
    emptrow[i] = myLineArray.length;//最后加上总行数

    //================文本图层变量设置================
    var docRef = app.activeDocument;//当前打开的文档
    var fontName = "MicrosoftYaHei";// 定义字体：微软雅黑
    var textColor = new SolidColor();//定义字体颜色：美漫黑=R33 G33 B33
    textColor.rgb.red = 33;
    textColor.rgb.green = 33;
    textColor.rgb.blue = 33;
    //读取文档宽度与高度
    var width = docRef.width;
    var height = docRef.height;

    var row_len = 0;
    for (var n = 0; n < i; n++) {
        row_len = emptrow[n + 1] - emptrow[n] - 1;//计算每个段落包含的行数

        var artLayerRef = docRef.artLayers.add();//添加图层
        artLayerRef.kind = LayerKind.TEXT;//转为文本图层
        var textItemRef = artLayerRef.textItem;

        /********************
         * 左对齐=LEFT
         * 右对齐=RIGHT
         * 居中对齐=CENTER
         ********************/
        textItemRef.justification = Justification.CENTER; //对齐方式
        /********************
         * 犀利=CRISP
         * 无=NONE
         * 锐利=SHARP
         * 平滑=SMOOTH
         * 浑厚=STRONG
         ********************/
        textItemRef.antiAliasMethod = AntiAlias.STRONG; //消除锯齿方式
        /********************
         * 手动指定=MANUAL
         * 度量标准=METRICS
         * 视觉=OPTICAL（推荐这个，字距更为紧凑）
         ********************/
        textItemRef.autoKerning = AutoKernType.OPTICAL; //字符间距微调
        textItemRef.color = textColor;//字体颜色
        textItemRef.size = 30; //字号
        textItemRef.font = fontName;
        textItemRef.position = Array(width * 0.1 + n * (width / (i + 1)), height * 0.1 + n * (height / (i + 1)));//位置
        textItemRef.contents = writeOnLayerS(n, row_len, myLineArray, emptrow);//写入图层
    }
//================清空变量================
    myLineArray = null;
    i = null;
    n = null;
    row_len = null;
    emptrow = null;
    docRef = null;
    textColor = null;
    newTextLayer = null;
}

//================设置区域================
// var read_method = '对话框选取';
var read_method = '读取特定文档';

//================启动程序================
main_process(read_method);