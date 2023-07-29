// 获取当前打开的文档
var doc = app.activeDocument;

// 如果当前文档不是多通道模式，图层合并为背景层进行

if (doc.mode != DocumentMode.GRAYSCALE) {
  // 检查当前图层是否为灰色层
	alert("文件格式不对！需要是已分色的灰度文件。");
 
	


} else {
 
	// 获取当前文档
var doc = app.activeDocument;
// 获取所有通道
var allChannels = doc.channels;
// 获取通道数量
var numChannels = allChannels.length;
// 勾选所有通道并使其可见
var activeChannels = [];
for (var i = 0; i < numChannels; i++) {
  var channel = allChannels[i];
  channel.visible = true;
  activeChannels.push(channel);
}
// 将勾选的通道设为当前文档的激活通道
doc.activeChannels = activeChannels;

	// 设置单位与标尺选项为像素
app.preferences.rulerUnits = Units.PIXELS; 
	
// 获取当前文档
	var doc = app.activeDocument;
	var width = doc.width.as('px');
	var height = doc.height.as('px');
// 获取文档
var doc = app.activeDocument;
// 获取文档分辨率
var resolution = doc.resolution;
// 如果分辨率不等于381，则将分辨率设置为381，并重新采样
if (resolution != 381) {
  doc.resizeImage(null, null, 381, ResampleMethod.NEARESTNEIGHBOR);
}
// 将单位设置为厘米
preferences.rulerUnits = Units.CM;
// 获取文档的宽度和高度
var width = doc.width;
var height = doc.height;
// 计算新的宽度和高度
var newWidth = width;
var newHeight = height + 5; // 加大5cm，上下各加 2.5cm
// 改变画布大小
doc.resizeCanvas(newWidth, newHeight, AnchorPosition.MIDDLECENTER);
// 在当前文档上新建一个图层
var newLayer = doc.artLayers.add();
newLayer.name = "Selection";
// 获取文档的宽度和高度
var width = doc.width;
var height = doc.height;

//--------------------框选横线,并填入黑色-----------------------
// 计算选区的位置和尺寸
var rectWidth = 1.5 / 2.54 * 381; // 选区宽度为 1.5cm，转为像素单位
var rectHeight = 0.1 / 2.54 * 381; // 选区高度为 0.1cm，转为像素单位
var canvasX = width / 2.54 * 381; //画布宽度一半
var canvasY = height / 2.54 * 381; // 画布高度一半

// 创建横线选区
var hxLeft = canvasX/2 - rectWidth / 2; // 选区左边距离文档左边框的距离
var hxTop = rectWidth / 2 - rectHeight / 2; //选区顶部(十字线横线)距离文档顶部的距离
var hxSelection = doc.selection;
hxSelection.select([
    [hxLeft, hxTop], // 左上角
    [hxLeft + rectWidth, hxTop], // 右上角
    [hxLeft + rectWidth, hxTop + rectHeight], // 右下角
    [hxLeft, hxTop + rectHeight], // 左下角
    [hxLeft, hxTop] // 回到左上角
], SelectionType.REPLACE, 0, false);  //不使用抗锯齿
// 填充选区
var fillColor = new SolidColor();
fillColor.rgb.hexValue = "000000"; // 黑色
doc.selection.fill(fillColor);
// 取消选区
doc.selection.deselect();
//--------------------------------------
//--------------------再框画小横线------
var hxdSelection = doc.selection;
hxdSelection.select([
    [canvasX/2 - rectWidth/4 , rectWidth/4 - rectHeight/4], // 左上角
    [canvasX/2 + rectWidth/4 , rectWidth/4 - rectHeight/4], // 右上角
    [canvasX/2 + rectWidth/4 , rectWidth/4 + rectHeight/4], // 右下角
    [canvasX/2 - rectWidth/4 , rectWidth/4 + rectHeight/4], // 左下角
    [canvasX/2 - rectWidth/4 , rectWidth/4 - rectHeight/4] // 回到左上角
], SelectionType.REPLACE, 0, false);  //不使用抗锯齿
// 填充选区
var fillColor = new SolidColor();
fillColor.rgb.hexValue = "000000"; // 黑色
doc.selection.fill(fillColor);

// 拷贝已有内容,粘贴下条
	doc.selection.copy();
var hxSelection = doc.selection;
hxSelection.select([
    [canvasX/2 - rectWidth/4 , rectWidth/4*3 - rectHeight/2], // 左上角
    [canvasX/2 + rectWidth/4 , rectWidth/4*3 - rectHeight/2], // 右上角
    [canvasX/2 + rectWidth/4 , rectWidth/4*3 + rectHeight/2], // 右下角
    [canvasX/2 - rectWidth/4 , rectWidth/4*3 + rectHeight/2], // 左下角
    [canvasX/2 - rectWidth/4 , rectWidth/4*3 - rectHeight/2] // 回到左上角
], SelectionType.REPLACE, 0, false);  //不使用抗锯齿
//粘贴
	doc.paste();
// 把第一个图层向下合并到下一个图层上
app.activeDocument.layers[0].merge();
// 取消选区
doc.selection.deselect();
//-------------------------------------------------------------------
// 选取原来的已有内容
	doc.selection.select([[hxLeft,0],[hxLeft+rectWidth,0],[hxLeft+rectWidth,rectWidth],[hxLeft,rectWidth]]);
// 拷贝已有内容
	doc.selection.copy();
	doc.paste(); 
	// 将新的图层旋转90度
	doc.activeLayer.rotate(90);
// 将所有可见图层合并为单个图像层
	doc.flatten();
/*
粘贴到新层后,如何让新层的白色内容改成透明
//--------------------框选一直线,并填入黑色-----------------------
// 创建选区
var sxLeft = canvasX / 2 - rectHeight / 2; // 左上角
var sxTop = canvasY - canvasY ; //选区顶部(十字线横线)距离文档顶部的距离
var sxSelection = doc.selection;
sxSelection.select([
    [sxLeft, sxTop], // 左上角
    [sxLeft + rectHeight, sxTop], // 右上角
    [sxLeft + rectHeight, rectWidth], // 右下角
    [sxLeft, rectWidth], // 左下角
    [sxLeft, sxTop] // 回到左上角
], SelectionType.REPLACE, 0, false);  //不使用抗锯齿
// 填充选区
var fillColor = new SolidColor();
fillColor.rgb.hexValue = "000000"; // 黑色
doc.selection.fill(fillColor);
// 取消选区
doc.selection.deselect();
//-------------------------------------------------------------------
*/

// 选取原来的已有内容
	doc.selection.select([[hxLeft,0],[hxLeft+rectWidth,0],[hxLeft+rectWidth,rectWidth],[hxLeft,rectWidth]]);

// 拷贝已有内容
	doc.selection.copy();
//选取位置
doc.selection.select([[hxLeft,canvasY-rectWidth],[hxLeft+rectWidth,canvasY-rectWidth],[hxLeft+rectWidth,canvasY],[hxLeft,canvasY]
], SelectionType.REPLACE, 0, false);//不使用抗锯齿
//粘贴
	doc.paste();
// 将所有可见图层合并为单个图像层
doc.flatten();
}