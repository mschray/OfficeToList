﻿// This function is called when Office.js is ready to start your Add-in
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

        //todo
        // people with only first name (leslie
        // people with middle initial (shaun S. Sharma)

        // some email address are wrapped in single quoutes (e.g. 'Brad Becker') which is causing messing up the formatting in the html when added to the page
        // this code is meant to address this, but should be encoding it instead?
        var displayName = data.value[i].displayName;
        if (data.value[i].displayName.indexOf("'") >= 0) {
            // do a global replace on string
            displayName = data.value[i].displayName.replace(/'/g, "");
        }
        
        //<li class="ui-widget-content">Item 1</li>
        //list = "<li class=\"ui-widget-content\">" + data.value[i].displayName + "</li>";
        list = "<input type='checkbox' name='displayName' value=' " + displayName + "' >" + displayName+"<br>";
        $("#display-list").append(list);
    }
}

var prependNames = function prependNames() {
    console.log("prependNames() called");
    var item = Office.context.mailbox.item;

    // hold the list of name in the this array
    var selectedName = [];
        
    $("input:checked").each(function () {
        var that = $(this);
        console.log("This is the checklist item :" + that[0].checked);
        console.log("This is the checklist item :" + that[0].value);

        // variable for the full name
        var fullName = that[0].value;
        
        // remove the (blah blah in name like MS vendors)
        // for example Andrew Baker (Piraeus Data LLC) <v-andbak@microsoft.com>
        // if it ends with a ) than do this
        // or 
        // had another example where the maden name was used in the middle of the display name
        // Shuhan (Bali) Majid <smajid@microsoft.com>

     /*   if (fullName[fullName.length - 1] == ')') {
            var startParen = fullName.lastIndexOf('(');
            if (startParen > 0) {
                fullName = fullName.slice(0, startParen);
                console.log("Remove of paren :" + fullName);
                fullName = fullName.trim();
                console.log(" of paren :" + fullName);
            }

        // had another example where the maden name was used in the middle of the display name
        // Shuhan (Bali) Majid <smajid@microsoft.com>
        } else */ 
        if (fullName.lastIndexOf('(')) {
            var startParen = fullName.lastIndexOf('(');
            var endParen = fullName.lastIndexOf(')');

            fullName = fullName.substring(1, startParen - 1) + fullName.substring(endParen, fullName.length - 1);
            fullName = fullName.trim();
        }

        // get the first name
        // I figure the name will have space between first and land name
        var locationOfLastSpace = fullName.lastIndexOf(' ');

        selectedName.push(fullName.slice(0,locationOfLastSpace));
    });

    // format the name
    var formattedNameStr ='';

    // its only one name selected
    if (selectedName.length == 1) {
        item.body.prependAsync(selectedName.join("") + ",\n");
    } else if (selectedName.length == 2) { // two names selected
        item.body.prependAsync(selectedName[0] + ' and ' + selectedName[1] + ",\n");
    } else if (selectedName.length > 2) {  // more that two names selected 

        for (var i = 0; i < selectedName.length - 1; i++) {
            formattedNameStr += selectedName[i];
            if (i < selectedName.length - 2) {
                formattedNameStr += ", ";
            }
        }

        formattedNameStr += ' and ' + selectedName[selectedName.length-1] +',';

        item.body.prependAsync(formattedNameStr + "\n");
    
    }

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