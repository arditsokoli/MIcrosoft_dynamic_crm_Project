var Sdk = window.Sdk || {};
(function () {
this.formOnChange = function (executionContext) {
    var formContext = executionContext.getFormContext();  
    var revenue = formContext.getAttribute("fisoft_revenue").getValue();
    var address = formContext.getAttribute("fisoft_address").getValue();
    if(parseInt(revenue)>1000){
    formContext.ui.setFormNotification("Top Customer","INFO","someid");
    }else{
        formContext.ui.clearFormNotification("someid");
    }
if((address.includes("Albania"))||(address.includes("albania"))||(address.includes("ALBANIA"))){
    formContext.getControl("fisoft_revenue").setVisible(true);
}else{
    formContext.getControl("fisoft_revenue").setVisible(false);
}

}

this.formOnLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var formtype = formContext.ui.getFormType();    
    var revenue = formContext.getAttribute("fisoft_revenue").getValue();
    var address = formContext.getAttribute("fisoft_address").getValue();
    if(formtype != 1){
    if((address.includes("Albania"))||(address.includes("albania"))||(address.includes("ALBANIA"))){
        formContext.getControl("fisoft_revenue").setVisible(true);
    }else{
        formContext.getControl("fisoft_revenue").setVisible(false);
    }
    if(parseInt(revenue)>1000){
    formContext.ui.setFormNotification("Top Customer","INFO","someid");
    }
}
}
}).call(Sdk);