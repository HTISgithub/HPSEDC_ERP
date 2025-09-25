var path;
var poly;
var directionsService;
var directionsDisplay;

var MapPoints = '';
var map;
var bounds;
var boundLatlng;

var markersG = new Array();

function initMap() {
    map = new google.maps.Map(document.getElementById('map'), {
        zoom: 10,
        center: { lat: 30.7145, lng: 76.7149 }
    });
    bounds = new google.maps.LatLngBounds();
    initPath();
    polyline();
}

function resetBounds() {
    bounds = new google.maps.LatLngBounds();
}

function initPath() {
    path = new google.maps.MVCArray();
    //bounds = new google.maps.LatLngBounds();
}

function polyline() {
    var lineSymbol = {
        path: google.maps.SymbolPath.FORWARD_CLOSED_ARROW
    };
    poly = new google.maps.Polyline({
        map: map, strokeColor: 'blue',
        icons: [{
            icon: lineSymbol,
            offset: '50%'
        }],
    });

    poly.setMap(map);
    poly.setPath(path);
}

function removePolyline() {
    poly.setMap(null);
}

function drawPath(MapPoints, showM) {
    //$('#ProgressImg').show();
    if (MapPoints.length > 0) {
        console.log('asdasd');
        var st = Math.ceil(MapPoints.length / 95);
        console.log(st);
        console.log(MapPoints);
        var coords = [];
        removePolyline();
        initPath();
        polyline();
        //var myLatlng;
        //var boundLatLng;
        var j = 0;
        for (var i = 0; i < MapPoints.length; i++) {
            if ((i + 1) < MapPoints.length) {
                boundLatlng = new google.maps.LatLng(MapPoints[i].lat, MapPoints[i].lng);
              
                path.push(boundLatlng);
                bounds.extend(boundLatlng);
                if (j < MapPoints.length - 1) {
                coords.push(MapPoints[j].lat + ',' + MapPoints[j].lng);
                }
                j = j + st;
            }
          
        }
       
        if (showM == "N") {
            setTimeout(mapBounds(), 1000);
        }
        coords.push(MapPoints[MapPoints.length - 1].lat + ',' + MapPoints[MapPoints.length - 1].lng)
        console.log('array: ', coords);
     
        console.log(distancePolyline());
        if ($('#spanKmsSys').length) {
            $('#spanKmsSys').html(distancePolyline());
        }
        if ($('#spanKms').length) {
            $('#spanKms').html(distancePolyline());
        }
    }
}

function distancePolyline() {
    return (google.maps.geometry.spherical.computeLength(poly.getPath()) / 1000).toFixed(2);

}


var snappedCoordinates = [];
var placeIdArray = [];
var finalArray = [];
var polylines = [];

function runSnapToRoad(finalArray) {
    $.get('https://roads.googleapis.com/v1/snapToRoads', {
        interpolate: true,
        key: 'AIzaSyDzeigeMmQ3yiZgb813di4VUe7c_fATyVw',
        path: finalArray.join('|')
    }, function (data) {
        processSnapToRoadResponse(data);
        drawSnappedPolyline();
    });
}

// Store snapped polyline returned by the snap-to-road service.
function processSnapToRoadResponse(data) {
    snappedCoordinates = [];
    placeIdArray = [];
    for (var i = 0; i < data.snappedPoints.length; i++) {
        var latlng = new google.maps.LatLng(
          data.snappedPoints[i].location.latitude,
          data.snappedPoints[i].location.longitude);
        snappedCoordinates.push(latlng);
        placeIdArray.push(data.snappedPoints[i].placeId);
    }
}

// Draws the snapped polyline (after processing snap-to-road response).
function drawSnappedPolyline() {
    var snappedPolyline = new google.maps.Polyline({
        path: snappedCoordinates,
        strokeColor: 'blue',
        strokeWeight: 3
    });
    snappedPolyline.setMap(map);
    removePolyline();
    polyline();
    poly.push(snappedPolyline);

}

function showMarker(mapPoints) {
    var startPoint;
    var endPoint;
    console.log('dcvsd');
    console.log(mapPoints);
    var json = mapPoints;
    startPoint = json[0].lat + ',' + json[0].lng;
      endPoint = json[json.length - 1].lat + ',' + json[json.length - 1].lng;
        //var bounds = new google.maps.LatLngBounds();
    //var myLatlng;

    wayPointsG = wayPointCreate(json, false);
     if ($('#map').length > 0) {
        var icon = "";
        for (i = 0; i < json.length; i++) {

            if (json[i].WorkStatus) {
                switch (json[i].WorkStatus) {
                    case "Started":
                        icon = 'http://admin.htistelecom.in/images/mapicon/start.png';
                        break;
                    case "Approved":
                        icon = 'http://admin.htistelecom.in/images/mapicon/approved.png';
                        break;
                    case "Completed":
                        icon = 'http://admin.htistelecom.in/images/mapicon/complete.png'
                        break;
                    case "Rejected":
                        icon = 'http://admin.htistelecom.in/images/mapicon/reject.png';
                        break;
                    case "Pending":
                        icon = 'http://admin.htistelecom.in/images/mapicon/pending.png';
                        break;
                    case "Assigned":
                        icon = 'http://admin.htistelecom.in/images/mapicon/start.png';
                        break;
                }
            }
            else {
                
                icon = 'http://admin.htistelecom.in/images/mapicon/pending.png';
               
            }
            boundLatlng = new google.maps.LatLng(json[i].lat, json[i].lng);
            
            bounds.extend(boundLatlng);
            var marker = new google.maps.Marker({
                position: new google.maps.LatLng(json[i].lat, json[i].lng),
                map: map,
                icon: new google.maps.MarkerImage(icon)
                //title: contenthtml(json[i].WorkDate, json[i].ProjectName, json[i].SiteName, json[i].ActivityText, json[i].WorkStatus, json[i].lat, json[i].lng)
            });
            markersG.push(marker);

            var loc = new google.maps.LatLng(marker.position.lat(), marker.position.lng());
            // bounds.extend(loc);
          
            marker['infowindow'] = new google.maps.InfoWindow({
                content: contenthtml(json[i].WorkDate, json[i].ProjectName, json[i].SiteName, json[i].ActivityText, json[i].WorkStatus, json[i].lat, json[i].lng, json[i].SiteUploadedId)

                //content: contenthtml(workDate, project, site, activity, status, lat, lng)
            });
          
            google.maps.event.addListener(marker, 'click', function () {
                //infowindow.open(map, this);
                this['infowindow'].open(map, this);
            });

            if (i == json.length - 1) {

            }
        }
        
        //map.fitBounds(bounds);
        setTimeout(mapBounds(), 1000);
    }

}
function wayPointCreate(json, stepover) {
    var waypts = [];
    for (var i = 0; i < json.length; i++) {
        waypts.push({
            location: json[i].lat + "," + json[i].lng,
            stopover: stepover
        });
    }
    return waypts;
}

function mapBounds() {
    map.fitBounds(bounds);
}

function clearMarkers() {
    for (var i = 0; i < markersG.length; i++) { 
        markersG[i].setMap(null);
    }
}
function showEmpMap(mapPoints) {
    var startPoint;
    var endPoint;
    console.log('dcvsd');
    console.log(mapPoints);
    var json = mapPoints;
     wayPointsG = wayPointCreate(json, false);
     if ($('#map').length > 0) {
        var icon = "";
        for (i = 0; i < json.length; i++) {
            icon = 'http://admin.htistelecom.in/images/mapicon/start.png';
            boundLatlng = new google.maps.LatLng(json[i].lat, json[i].lng);
            bounds.extend(boundLatlng);
            var marker = new google.maps.Marker({
                position: new google.maps.LatLng(json[i].lat, json[i].lng),
                map: map,
                icon: new google.maps.MarkerImage(icon)
                });
            markersG.push(marker);

            var loc = new google.maps.LatLng(marker.position.lat(), marker.position.lng());
            marker['infowindow'] = new google.maps.InfoWindow({
                content: EmpInfoHtml(json[i].WorkDate, json[i].EmpCode, json[i].EmpName, json[i].Designation, json[i].Mobile, json[i].lat, json[i].lng, json[i].EmpID)
                              
            });
            google.maps.event.addListener(marker, 'click', function () {
                      this['infowindow'].open(map, this);
            });

            if (i == json.length - 1) {

            }
        }
        setTimeout(mapBounds(), 1000);
    }

}