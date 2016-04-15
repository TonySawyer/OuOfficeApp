/// <reference path="../App.js" />

(function () {
    "use strict";
    var userDetails = {
        PI: "E2762956",
        Name: "Ian Brown",
        };

    var courses = {
        Course:  {
            Title: "L192 Beginners French",
            Code: "L192",
            TutorName: "Mrs Tutor",
            TutorContactEmail: "mrstutor@open.ac.uk",
            TutorContactVoip: "mrstutor@open.ac.uk",
            TMAS: [
                {Title: "TMA01", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764144"},
                {Title: "TMA02", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764174"},
                {Title: "TMA03", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764189"},
                {Title: "TMA04", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764367"},
                {Title: "EMA", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764196"}
            ]
        }
    }

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#insert-standard-header').click(insertStandardHeader);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function insertStandardHeader() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections;

            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');

            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Create a proxy object the primary header of the first section. 
                // Note that the header is a body object.
                var myHeader = mySections.items[0].getHeader("primary");

                // Queue a command to insert text at the end of the header.
                myHeader.insertText(getHeaderText(), Word.InsertLocation.end);

                // Queue a command to wrap the header in a content control.
                myHeader.insertContentControl();

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log("Added a header to the first section.");
                });
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });



        //var context = Office.context;
        //var document = context.document;
        //var sections = document.sections;
        //context.load(sections, 'body/style');
        //context.sync().then(function () {
        //    var header = sections.items[0].getHeader("primary");
        //    header.insertText(getHeaderText(), Word.InsertLocation.end);
        //    header.insertContentControl();
        //    contex.sync();

        //});

        //displayAllBindings();
        //if (Office.context.document.setSelectedDataAsync) {
        //    //Office.context.document.goToByIdAsync
            
        //    Office.context.document.setSelectedDataAsync(getHeaderText(), function (result) {
        //        //Upon return, if the call was unable to insert text, let the user know.
        //        if (result.status === Office.AsyncResultStatus.Failed) {
        //            app.showNotification("There's a problem!", "The sample text was unable to be inserted.");
        //        }
        //    });
        //} else {
        //    app.showNotification("There's a problem!", "This product does not support inserting content.");
        //}
    }

    function displayAllBindings() {
        write("getting bindings");
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            var bindingString = '';
            for (var i in asyncResult.value) {
                bindingString += asyncResult.value[i].id + '\n';
            }
            write('Existing bindings: ' + bindingString);
        });
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }

    function getHeaderText() {
        return "Ian Brown (PI E2762956) L192 ETMA01";
    }
})();