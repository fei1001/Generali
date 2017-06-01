// NOTE: This code is unsupported since it is not using the SDK
// This is the only way of exactly mimicking the behaviour of the out of the box 
// dialog windows that are reliably sized and positioned
(function () {
    Develop1_RibbonCommands_runDialogGrid = function (ids, objectTypeCode, dialogId) {
        if ((ids == null) || (!ids.length)) {
            alert(window.LOCID_ACTION_NOITEMSELECTED);
            return;
        }
        if (ids.length > 1) {
            alert(window.LOCID_GRID_TOO_MANY_RECORDS_IWF);
            return;
        }
        var rundialog = Mscrm.CrmUri.create('/cs/dialog/rundialog.aspx');
        rundialog.get_query()['DialogId'] = dialogId;
        rundialog.get_query()['ObjectId'] = ids[0];
        rundialog.get_query()['EntityName'] = objectTypeCode;
        openStdWin(rundialog, buildWinName(null), 615, 480, null);
    }
    Develop1_RibbonCommands_runDialogForm = function (objectTypeCode, dialogId) {
        var primaryEntityId = Xrm.Page.data.entity.getId();
        var rundialog = Mscrm.CrmUri.create('/cs/dialog/rundialog.aspx');
        rundialog.get_query()['DialogId'] = dialogId;
        rundialog.get_query()['ObjectId'] = primaryEntityId;
        rundialog.get_query()['EntityName'] = objectTypeCode;
        var hostWindow = window;
        if (typeof (openStdWin) == 'undefined') {
            hostWindow = window.parent; // Support for Turbo-forms in CRM2015 Update 1
        }
        if (typeof (hostWindow.openStdWin) != 'undefined') {
            hostWindow.openStdWin(rundialog, hostWindow.buildWinName(null), 615, 480, null);
        }
    }
})();