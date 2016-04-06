// This function is called when Office.js is ready to start your Add-in
Office.initialize = function (reason) {
    $(document).ready(function () {

        $("button:first" ).click(prependNames);



        addToLineRecipients();

    });
};


//var getBody = function (data) {
//    console.log("get body called");

//    var item = Office.context.mailbox.item;

//    item.body.prependAsync("Ryan");
//    //item.body.setAsync("Ryan\n\n" + data.value);
//}

var list;

// Adds the current user to the recipient list
var reciptList = function (data) {
    console.log("the recip list callback was reached");
    //	console.log("To data for callback:" + data.value.length);
    //	console.log("To data for callback:" + data.value[0]);
    console.log("To data for callback:" + data.value[0].displayName);

    
    for (var i = 0; i < data.value.length; i++) {
        console.log("Display Name: " + data.value[i].displayName);

        //<li class="ui-widget-content">Item 1</li>
        //list = "<li class=\"ui-widget-content\">" + data.value[i].displayName + "</li>";
        list = "<input type='checkbox' name='displayName' value=' " + data.value[i].displayName + "' >" + data.value[i].displayName+"<br>";
        $("#display-list").append(list);
    }

    //$('#list-recipients').text("etst");

    
    //item.body.getAsync(getBody);
    
}

var prependNames = function prependNames() {
    console.log("prependNames() called");
    var item = Office.context.mailbox.item;

    var selectedName = [];

    //var displayChecked = function () {

    //    var listOfControls = [];

    //    $("input:checked").each(function () {
    //        listOfControls.push($(this)[0].value);
    //    })

    //    $("#list").text(listOfControls.join(" ,"))

    //};

    $("input:checked").each(function () {
        var that = $(this);
        console.log("This is the checklist item :" + that[0].checked);
        console.log("This is the checklist item :" + that[0].value);
        selectedName.push($(this)[0].value);
    });

    item.body.prependAsync(selectedName.join(", ")+"\n");

}

function addToLineRecipients() {
    var item = Office.context.mailbox.item;

    // I don't need address to add since I am not inject myself into the to line anymore.  No need the reply or new mail adds me already.
    //var addressToAdd = {
    //    displayName: Office.context.mailbox.userProfile.displayName,
    //    emailAddress: Office.context.mailbox.userProfile.emailAddress
    //};

    // Make sure it a message item and if so add the list of recipants on the to line
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
        try {
            //var list = item.Entities.contacts();
            var list = item.to.getAsync(reciptList);

        } catch (msg) {
            console.log("Martin pluggin " + msg);
        }

        //Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        //Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
}