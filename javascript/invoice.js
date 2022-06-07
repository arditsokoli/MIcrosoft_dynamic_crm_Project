var Sdk = window.Sdk || {};
(function () {
this.formOnChange = function (executionContext) {
    var formContext = executionContext.getFormContext();    
	var paid = formContext.getAttribute("fisoft_paid").getValue();
	var number = formContext.getAttribute("fisoft_number").getValue();
    var createdDate = formContext.getAttribute("fisoft_createddate").getValue();
    //874630000 yes
    //874630001 no
if(parseInt(paid)==874630001){
 formContext.ui.setFormNotification("This invoice \""+number+"\" need Attention","WARNING","idv");
}else{
formContext.ui.clearFormNotification("idv");
Xrm.Page.getControl("fisoft_paid").setDisabled(true);
formContext.getAttribute("fisoft_paiddate").setValue(createdDate);
}
}

this.dateOnChange = function (executionContext) {
var formContext = executionContext.getFormContext(); 
var paid = formContext.getAttribute("fisoft_paid").getValue();
var createdDate = formContext.getAttribute("fisoft_createddate").getValue();
if(parseInt(paid)==874630000){
formContext.getAttribute("fisoft_paiddate").setValue(createdDate);
}
}

this.formOnLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();    
    var formtype = formContext.ui.getFormType();
    var paid = formContext.getAttribute("fisoft_paid").getValue();
    var number = formContext.getAttribute("fisoft_number").getValue();
    if((formtype == 2)&&(parseInt(paid)==874630001)){
        formContext.ui.setFormNotification("This invoice \""+number+"\" need Attention","WARNING","idv");
       }else{
        Xrm.Page.getControl("fisoft_paid").setDisabled(true); 
       }
    if(formtype == 1){
    var currentDate = new Date();
    formContext.getAttribute("fisoft_createddate").setValue(currentDate);
    formContext.getAttribute("fisoft_paid").setValue(874630001);  
    formContext.ui.setFormNotification("This new invoice need Attention","WARNING","idv");
    }
}
}).call(Sdk);
