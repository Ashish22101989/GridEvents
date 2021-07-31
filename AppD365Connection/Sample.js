
// A namespace defined for the sample code
// As a best practice, you should always define 
// a unique namespace for your libraries
var Oba = window.Oba || {};
Oba.Common = Oba.Common || {};
Oba.Appointment = Oba.Appointment || {};

(function () {
    // Define some global variables
    var myUniqueId = "_myUniqueId"; // Define an ID for the notification
    var currentUserName = Xrm.Utility.getGlobalContext().userSettings.userName; // get current user name
    var message = currentUserName + ": Your JavaScript code in action!";

    this.getFormContext = function (executionContext) {
        return executionContext.getFormContext();
    }

    this.getAttribute = function (executionContext, attributeName) {
        var formContext = this.getFormContext(executionContext);
        return formContext.getAttribute(attributeName);
    }

    // Code to run in the form OnLoad event
    this.setValue = function (executionContext, attributeName, value) {
        var formContext = this.getFormContext(executionContext);
        var attr = this.getAttribute(executionContext, attributeName);
        if (attr != null && attr != undefined) {
            attr.setValue(value);
            Xrm.Navigation.openAlertDialog({ text: "value set saved." });
        }
    }

    // Code to run in the form OnLoad event
    this.getValue = function (executionContext, attributeName) {
        var formContext = this.getFormContext(executionContext);
        var attr = this.getAttribute(executionContext, attributeName);
        if (attr != null && attr != undefined) {
            Xrm.Navigation.openAlertDialog({ text: "get value called" });
            return attr.getValue();
        }
    }

}).call(Oba.Commmon);


(function () {

    // Code to run in the form OnLoad event
    this.formOnLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();
    }

    // Code to run in the column OnChange event 
    this.attributeOnChange = function (executionContext) {
        var formContext = executionContext.getFormContext();
    }

    // Code to run in the form OnSave event 
    this.formOnSave = function () {
        // Display an alert dialog
        Xrm.Navigation.openAlertDialog({ text: "Record saved." });
    }

}).call(Oba.Appointment);