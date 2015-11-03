// This function is called when Office.js is ready to start your Add-in
Office.initialize = function (reason) {
    $(document).ready(function () {

        $('#add-recipients').submit(prependNames());

        addToRecipients();

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
        list = data.value[i].displayName + "<br>";
    }

    //$('#list-recipients').text("etst");
    $("#display-list").append(list);
    var item = Office.context.mailbox.item;
    //item.body.getAsync(getBody);
    
}

function prependNames() {
    console.log("prependNames() called");

    item.body.prependAsync(list);

}

function addToRecipients() {
    var item = Office.context.mailbox.item;
    var addressToAdd = {
        displayName: Office.context.mailbox.userProfile.displayName,
        emailAddress: Office.context.mailbox.userProfile.emailAddress
    };

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
        try {
            //var list = item.Entities.contacts();
            var list = item.to.getAsync(reciptList);

        } catch (msg) {
            console.log("Martin pluggin " + msg);
        }
        Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
}