
layer.msg("haha")

var upload = layui.upload;

//选择当前文件
upload.render({
 elem: '#chooseNowBtn' //绑定元素
 ,url: '' //上传接口
 ,auto: false
 ,done: function(res){
  //上传完毕回调
 }
 ,error: function(){
 //请求异常回调
 layer.msg("fuck")
 }
});

//选择销售文件
//upload.render({
// elem: '#chooseSaleBtn' //绑定元素
// ,url: '/excel/uploadSale' //上传接口
// ,accept: "file"
// ,acceptMime: "file/xlsx"
// ,exts: "xlsx"
// ,auto: true
// ,done: function(res){
// layer.msg("nihc")
// console.log(res)
//  //上传完毕回调
// }
// ,choose: function(obj){
//
// }
// ,error: function(){
// //请求异常回调
// layer.msg("fuck")
// }
//});



