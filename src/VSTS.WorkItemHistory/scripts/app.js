VSS.require(["VSS/Service", "TFS/WorkItemTracking/RestClient", "VSS/Controls", "VSS/Controls/Grids", "TFS/WorkItemTracking/Services", "VSS/Controls/StatusIndicator"], function (VSS_Service, TFS_Wit_WebApi, Controls, Grids, workItemServices, StatusIndicator) {
    var projectId = VSS.getWebContext().project.id;
    var host = VSS.getWebContext().account.uri;
    var projectName = VSS.getWebContext().project.name;

    var isVSTS = VSS.getWebContext().account.uri.indexOf("visualstudio.com") > -1;
   
    var container = $("#data-container");
    var grid;
    var gridData;

    var witClient = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);

    var query = {
        query: "Select [System.ID] From WorkItems Where [System.TeamProject] = '" + projectName + "' Order by [System.ChangedDate] Desc"
    };

    var WITsHash = {};

    if (isVSTS) {
        //Currently TFS doesn't support the "color" property for the WITs
        witClient.getWorkItemTypes(projectName).then(function (result) {
            result.forEach(function (wit) { WITsHash[wit.name] = wit.color });
        });
    }

    witClient.queryByWiql(query).then(function (result) {
        // Generate an array of all work item ID's
        var historyWorkItems = result.workItems.map(function (wi) { return wi.id }).slice(0, 20);

        if (historyWorkItems.length <= 0)
        {
            $("#data-container").toggle();
            $("#btnLoadMore").toggle();
            $("#nothingToShow").toggle();
            VSS.notifyLoadSucceeded();
            return;
        }

        if (historyWorkItems.length < 20)
        {
            //hide the "Load More" button if the total items number is below the min
            $("#btnLoadMore").toggle();
        }

        var fields =[
            "System.Title",
            "System.State",
            "System.CreatedDate",
            "System.ChangedDate",
            "System.ChangedBy",
            "System.WorkItemType"];

        witClient.getWorkItems(historyWorkItems, fields).then(
            function (workItems) {
                // Access the work items and their field values

                gridData = workItems.map(function (w) {
                        return [
                            w.id,
                            w.fields["System.Title"],
                            w.fields["System.State"],
                            w.fields["System.ChangedDate"],
                            w.fields["System.ChangedBy"],
                            w.fields["System.WorkItemType"]];
                            });

                // We now have the open work items with field values, time to display these
                var options = {
                    width: "100%",
                    height: "90%",
                    useBowtieStyle: true,
                    source: gridData,
                    sortOrder: [
                        {
                            index: 3,
                            order: "desc"
                        }
                    ],
                    autoSort: true,
                    columns: [
                        { text: "ID", index: 0, width: 50, canSortBy: false },
                        {
                            text: "",
                            width: 5,
                            index: 5,
                            canSortBy: false,                            
                            getCellContents: function (
                                rowInfo,
                                dataIndex,
                                expandedState,
                                level,
                                column,
                                indentIndex,
                                columnOrder) {

                                var workItemColor = "rgb(100, 100, 100)"; //Generic

                                if (isVSTS && Object.keys(WITsHash).length > 0) {
                                    workItemColor = '#' + WITsHash[this.getColumnValue(dataIndex, column.index)];
                                }
                                else {
                                    switch (this.getColumnValue(dataIndex, column.index)) {
                                        case "Epic":
                                            workItemColor = "rgb(225, 123, 0)"; //Epic
                                            break;
                                        case "Feature":
                                            workItemColor = "rgb(119, 59, 147)"; //Feature
                                            break;
                                        case "Impediment":
                                            workItemColor = "rgb(255, 157, 0)"; //Impediment
                                            break;
                                        case "Product Backlog Item":
                                            workItemColor = "rgb(0, 156, 204)"; //PBI
                                            break;
                                        case "Task":
                                            workItemColor = "rgb(242, 203, 29)"; //Task
                                            break;
                                        case "Test Case":
                                            workItemColor = "rgb(255, 157, 0)"; //Test Case
                                            break;
                                        case "Bug":
                                            workItemColor = "rgb(204, 41, 61)"; //Bug
                                            break;
                                        default:
                                            workItemColor = "rgb(100, 100, 100)"; //Generic
                                    }
                                }                                
                               
                                return $("<div class='grid-cell'/>")
                                    .width(column.width || 5)
                                    .css("background-color", workItemColor)
                                    .css("background-clip", "border-box")
                                    .css("margin-top", "5px");
                            }
                        },
                        {
                            text: "Type",
                            index: 5,
                            width: 150,
                            canSortBy: false,
                            getCellContents: function (
                                rowInfo,
                                dataIndex,
                                expandedState,
                                level,
                                column,
                                indentIndex,
                                columnOrder) {

                                var item = grid.getRowData(dataIndex);                                

                                return $("<div class='grid-cell'/>")
                                    .width(column.width || 150)
                                    .text(this.getColumnValue(dataIndex, column.index));
                            }
                        },
                        {
                            text: "Title",
                            index: 1,
                            width: 400,
                            canSortBy: false,
                            getCellContents: function (
                                rowInfo,
                                dataIndex,
                                expandedState,
                                level,
                                column,
                                indentIndex,
                                columnOrder) {

                                var item = grid.getRowData(dataIndex);                                

                                return $("<div class='grid-cell'/>")
                                    .width(column.width || 400)
                                    .hover(function () { $(this).css("text-decoration", "underline") }, function () { $(this).css("text-decoration", "none") })
                                    .css("cursor", "pointer")
                                    .click(function () {
                                        workItemServices.WorkItemFormNavigationService.getService().then(function (workItemNavSvc) {
                                            workItemNavSvc.openWorkItem(item[0]);
                                        });
                                    })
                                    .text(this.getColumnValue(dataIndex, column.index));
                            }
                        },
                        { text: "Current State", index: 2, width: 100, canSortBy: false },
                        {
                            text: "Last Changed",
                            index: 3,
                            width: 175,
                            getCellContents: function (
                                rowInfo,
                                dataIndex,
                                expandedState,
                                level,
                                column,
                                indentIndex,
                                columnOrder) {

                                // Calculates the difference between current time and ChangeDate field value in days
                                var oneDay = 24 * 60 * 60 * 1000;
                                var today = new Date();
                                var changeDate = new Date(this.getColumnValue(dataIndex, column.index));
                                var diffDays = Math.round(Math.abs((changeDate.getTime() - today.getTime()) / (oneDay)));

                                var text = $.timeago(changeDate);
                                // If more than 4 days has passed, show in orange
                                // If more than 10 days has passed, show in red
                                return $("<div class='grid-cell'/>")
                                    .width(column.width || 100)
                                    .css("color", diffDays < 4 ? "black" : (diffDays < 10 ? "orange" : "red"))
                                    .text(text);
                            }
                        },
                        {
                            text: "Last Changed By",
                            index: 4,
                            width: 200,
                            canSortBy: false,
                            getCellContents: function (
                                rowInfo,
                                dataIndex,
                                expandedState,
                                level,
                                column,
                                indentIndex,
                                columnOrder) {

                                var text = this.getColumnValue(dataIndex, column.index);
                                var email = text.substring(text.indexOf("<") + 1, text.indexOf(">"));

                                var imageString = host + "/_api/_common/IdentityImage?id=&identifier=" + email.replace("@", "%40") + "&resolveAmbiguous=false&identifierType=0&size=0&__v=5";

                                return $("<div class='grid-cell'/>")
                                    .width(column.width || 100)
                                    .css("background", "url('" + imageString + "') no-repeat left")
                                    .css("padding-left", "40px")
                                    .text(text.substring(0, text.indexOf("<")));
                            }
                        }
                    ]
                };
                
                grid = Controls.create(Grids.Grid, container, options);

                VSS.notifyLoadSucceeded();
            });
    });

    var waitControlOptions = {
        target: $("#waitTarget"),
        cancellable: true
    };

    var waitControl = Controls.create(StatusIndicator.WaitControl, container, waitControlOptions);

    $("#btnLoadMore").click(function () {
        waitControl.startWait();

        witClient.queryByWiql(query).then(function (result) {
            // Generate an array of all open work item ID's
            var historyWorkItems = result.workItems.map(function (wi) { return wi.id }).slice(gridData.length, gridData.length + 20);

            if (historyWorkItems.length < 20) {
                //hide the "Load More" button if the total items number is below the min
                $("#btnLoadMore").toggle();
            }

            var fields = [
            "System.Title",
            "System.State",
            "System.CreatedDate",
            "System.ChangedDate",
            "System.ChangedBy",
            "System.WorkItemType"];

            witClient.getWorkItems(historyWorkItems, fields).then(
                function (workItems) {
                    var source = workItems.map(function (w) {
                        return [
                            w.id,
                            w.fields["System.Title"],
                            w.fields["System.State"],
                            w.fields["System.ChangedDate"],
                            w.fields["System.ChangedBy"],
                            w.fields["System.WorkItemType"]];
                    });
                    gridData = gridData.concat(source);
                    grid.setDataSource(gridData);

                    waitControl.endWait();
            });
        });        
    });
    
});


