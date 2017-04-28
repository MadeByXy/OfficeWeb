<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Setting.aspx.cs" Inherits="OfficeWeb.Setting" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Office在线设置页面</title>
    <link type="text/css" rel="stylesheet" href="/Resources/Style/bootstrap/css/bootstrap.min.css" />
    <link type="text/css" rel="stylesheet" href="/Resources/Style/bootstrap/css/docs.min.css" />
    <link type="text/css" rel="stylesheet" href="/Resources/Style/Setting.css" />
</head>
<body>
    <div ng-app="app" ng-controller="controller">
        <div class="panel panel-info">
            <div class="panel-heading">
                <h3 class="panel-title">设置页</h3>
            </div>
            <div ng-cloak class="panel-body ng-cloak">
                <div class="input-group">
                    <span class="input-group-addon">设置文件过期时间（天）</span>
                    <input type="text" class="form-control" ng-model="clearDay">
                </div>
            </div>
            <div ng-cloak class="panel-body ng-cloak">
                <div class="btn-group btn-group-justified">
                    <div class="btn-group" role="group">
                        <button type="button" class="btn btn-default" ng-click="Query('Function', 'Clear', '', '清理成功')">清理过期缓存文件</button>
                    </div>
                    <div class="btn-group" role="group">
                        <button type="button" class="btn btn-default btn-center" ng-click="Query('Function', 'ClearAll', '', '清理成功')">清理所有缓存文件</button>
                    </div>
                    <div class="btn-group" role="group">
                        <button type="button" class="btn btn-default" ng-class="{true:'active'}[auto]" ng-click="AutoClear()">自动清理过期文件</button>
                    </div>
                </div>
            </div>
        </div>
        <div class="panel panel-info">
            <div class="panel-heading">
                <h3 class="panel-title">COM检测</h3>
            </div>
            <div ng-cloak class="panel-body ng-cloak">
                <p>COM组件可以使OFFICE在页面总呈现更好的效果。</p>
                <div class="bs-callout bs-callout-info" ng-repeat="item in environment">
                    <h4>{{item.name}}检测</h4>
                    <div class="alert" ng-repeat="node in item.data" ng-class="{true:'alert-success',false:'alert-danger'}[node.value]">
                        <span ng-bind="node.name"></span>
                        <span ng-show="node.value">检测成功</span>
                        <span ng-show="!node.value">检测失败</span>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <script type='text/javascript' src='/Resources/Scripts/jquery-1.8.0.min.js'></script>
    <script type='text/javascript' src='/Resources/Scripts/angular.min.js'></script>
    <script type="text/javascript">
        angular.module('app', []).controller('controller', function($scope, $http) {
            $scope.auto = <%=OfficeWeb.Core.FileManagement.AutoClear%>;
            $scope.clearDay = <%=OfficeWeb.Core.FileManagement.FileClearDay%>;
            $scope.environment =<%=EnvironmentCheck%>;

            $scope.Query = function(action, name, value, message) {
                $http.get("/Setting.aspx?Action=" + action + "&name=" + name + "&value=" + value).success(function(){
                    alert(message||'设置成功');
                }).error(function(data){
                    alert(data);
                });
            }

            $scope.AutoClear = function() {
                $scope.auto = !$scope.auto;
                $scope.Query('Property','AutoClear',$scope.auto);
            }
        });
    </script>
</body>
</html>
