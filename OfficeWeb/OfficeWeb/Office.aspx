<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Office.aspx.cs" Inherits="OfficeWeb.Office" %>

<!DOCTYPE html>
<html>
<head runat="server">
    <title>Office在线预览</title>
    <link type="text/css" rel="stylesheet" href="/Resources/Style/Office.css" />
</head>
<body>
    <div ng-app="app" ng-controller="controller">
        <div ng-cloak class="Loading ng-cloak" ng-show="imageList == null || imageList.length == 0">
            <img alt="" src="/Resources/Images/Loading.gif" />
        </div>
        <div ng-cloak class="Error ng-cloak" ng-show="status == 'error'">
            <h1>:(</h1>
            <p class="Explain">真抱歉，页面发生了错误。</p>
            <p class="Hint"><span>错误原因：</span><span ng-bind="message"></span></p>
        </div>
        <div ng-cloak class="ImageList ng-cloak" ng-show="status == 'running' || status == 'finish'">
            <div class="Image" ng-repeat="item in imageList | orderBy : 'pageNum'">
                <img ng-src="{{item.imageUrl}}" />
            </div>
        </div>
        <div ng-cloak class="Loading_Bottom ng-cloak" ng-show="status == 'running' && imageList != null && imageList.length != 0">
            <img alt="" src="/Resources/Images/Loading.gif" />
        </div>
    </div>
    <script type='text/javascript' src='/Resources/Scripts/jquery-1.8.0.min.js'></script>
    <script type='text/javascript' src='/Resources/Scripts/angular.min.js'></script>
    <script type="text/javascript">
        angular.module('app', []).controller('controller', function ($scope, $http) {
            $scope.status = "loading";

            $scope.GetImage = function () {
                $http({
                    method: 'POST',
                    url: 'Office.aspx',
                    data: "Mode=GetResult&PageId=<%=SessionId%>",
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                }).success(function (data) {
                    $scope.status = data.status;
                    if (data.status == "running") {
                        $scope.imageList = data.imageList;
                        setTimeout(function () {
                            $scope.GetImage();
                        }, 100);
                    } else if (data.status == "finish") {
                        $scope.imageList = data.imageList;
                    } else {
                        $scope.message = data.message;
                    }
                });
            }

            $scope.GetImage();
        });
    </script>
</body>
</html>
