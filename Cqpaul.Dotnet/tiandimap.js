// JavaScript source code
var polygonMap;
var picMap;
var zoom = 12;
var img;
function loadMap() {
    //初始化地图对象
    $("#polygonDiv").show();
    $("#picDiv").hide();
    if (polygonMap == null) {
        polygonMap = new T.Map("polygonDiv");
    }
    //设置显示地图的中心点和级别
    polygonMap.centerAndZoom(new T.LngLat(106.55, 29.57), zoom);
    var points = [];
    points.push(new T.LngLat(106.31, 29.16));
    points.push(new T.LngLat(106.42, 29.37));
    points.push(new T.LngLat(106.53, 29.78));
    points.push(new T.LngLat(106.64, 29.99));
    //创建面对象
    var polygon = new T.Polygon(points, {
        color: "blue", weight: 3, opacity: 0.5, fillColor: "#FFFFFF", fillOpacity: 0.5
    });
    //向地图上添加面
    polygonMap.addOverLay(polygon);
}

function loadMapPic() {
    $("#picDiv").show();
    $("#polygonDiv").hide();
    if (picMap == null) {
        picMap = new T.Map('picDiv');
    }
    picMap.centerAndZoom(new T.LngLat(116.390750, 39.916980), zoom);
    var bd = new T.LngLatBounds(
        new T.LngLat(116.385360, 39.911380),
        new T.LngLat(116.395940, 39.921400));
    img = new T.ImageOverlay("http://lbs.tianditu.gov.cn/images/openlibrary/gugong.jpg", bd, {
        opacity: 1,
        alt: "故宫博物院"
    });
    picMap.addOverLay(img);
}