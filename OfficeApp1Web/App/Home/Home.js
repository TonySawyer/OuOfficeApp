/// <reference path="../App.js" />

(function () {
    "use strict";
    var userDetails = {
        PI: "E2762956",
        Name: "Ian Brown",
    };

    var courses = [
        {
            Title: "L192 Beginners French",
            Code: "L192",
            TutorDetails : {
                Name : "Mrs Tutor",
                Email: "https://msds.open.ac.uk/students/contacttutor.aspx?id=01700207&c=L192",
                Voip:"mrstutor@open.ac.uk"
            },
            Tmas: [
                { Title: "TMA01", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764144", WordCountRequired:"250"},
                { Title: "TMA02", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764174", WordCountRequired: "250" },
                { Title: "TMA03", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764189", WordCountRequired: "250" },
                { Title: "TMA04", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764367", WordCountRequired: "250" },
                { Title: "EMA", Url: "https://learn2.open.ac.uk/mod/oucontent/view.php?id=764196", WordCountRequired: "500" }
            ]
        }
    ];

    var currentCourseCode = "";
    var currentTMA = "";

    var serverUrl = "http://innovdata.azurewebsites.net/api/etmadata/User";
    var coursesServerUrl = "http://innovdata.azurewebsites.net/api/etmadata/Courses";


    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#connectOU').click(connectToOU);
            $('#insert-standard-header').click(insertStandardHeader);
            $('#submitETMA').click(submitETMA);
            $('#okSubmit').click(sendSubmission);
            $('#cancelSubmit').click(cancelSubmission);

            $("#moduleSelector").change(selectedCourseChanged);
            $("#tmaSelector").change(selectedTMAChanged);

            $('#chk1').change(setSubmitButtonEnabled);
            $('#chkCorrectFormat').change(setSubmitButtonEnabled);
            $('#chkNoCopying').change(setSubmitButtonEnabled);

        });
    };

    function connectToOU() {
        var userID = $('#studentId').val();
        var password = $('#password').val();
        if (userID == '' || password == '') {
            write('please provide your username and password.');
        }
        else {
            var result = getUserDetails(userID, password);
        }
    }

    function hideSpinner() {
        hide($('#spinner'));
    }

    function showSpinner() {
        show($('#spinner'));
    }

    function hide(divtoHide) {
        divtoHide.removeClass('shownPanel').addClass('hiddenPanel');
    }

    function show(divToShow) {
        divToShow.removeClass('hiddenPanel').addClass('shownPanel');

    }

    function getUserDetails(username, password) {
        hide($('#credentials'));
        hide($('#profile'));
        showSpinner();
        $.support.cors = true;
        $.ajax({
            url: serverUrl,
            type: 'GET',
            contentType: 'application/json;charset=utf-8'

        })
        .done(function (data) {
            displayUserDetails(data);
        })
        .fail(function (jqXHR, textStatus) {
            show($('#credentials'));
            write(jqXHR.statusText);
        })
        .always(function () {
            hide($('#spinner'));
        });
    }

    function getCoursesForUser(userId) {
        showSpinner();
        $.support.cors = true;
        $.ajax({
            url: coursesServerUrl,
            type: 'GET',
            contentType: 'application/json;charset=utf-8'

        })
        .done(function (data) {
            populateCourseDetails(data);
        })
        .fail(function (jqXHR, textStatus) {
            write(jqXHR.statusText);
        })
        .always(function () {
            hide($('#spinner'));
        });
    }

    function displayUserDetails(details) {
        show($('#profile'));
        show($('#mainPanels'));
        write('');
        $('#studId').text('Student ID: ' + details.Pi);
        $('#studName').text('Name: ' + details.Name);

        getCoursesForUser(details.Pi);
    }

    function setTutorContactEmailLink(tutorDetails) {

        var link = tutorDetails.EmailAddress;
        console.log(link);
        $('#contactTutor').text('Contact your tutor - ' + tutorDetails.Name);
        $("#contactTutor").attr("href", link);
    }


    function contactTutorClicked() {
        var data = $("#moduleSelector").val();
        var currentCourse = JSON.parse(data);
        var tutorEmailAddress = currentCourse.TutorContactEmail;
        var link = "mailto:" + tutorEmailAddress + ";subject=Contact from OU Student " + userDetails.Name;
        console.log(link);
        
        $("#mailToLink").attr("href", link);
        $("#mailToLink").trigger('click');

        window.open(link);
    }

    function populateCourseDetails(courseDetails) {

        $("#moduleSelector").empty();
        $("#tmaSelector").empty();

        $(courseDetails).each(function () {
            var item = $("<option />", {
                val: JSON.stringify(this),
                text: this.Title
            });
            item.appendTo($("#moduleSelector"));
        });

        selectCourse(courseDetails[0]);

    }

    function selectedCourseChanged() {
        var selectedCourse = $("#moduleSelector").val();
        selectCourse(selectedCourse);
    }

    function selectCourse(selectedCourse){
        $("#tmaSelector").empty();

        $(selectedCourse.Tmas).each(function () {
            var item = $("<option />", {
                val: JSON.stringify(this),
                text: this.Title
            });
            item.appendTo($("#tmaSelector"));
        });

        selectTMA(selectedCourse.Tmas[0]);
    }

    function selectedTMAChanged() {
        var selectedTMA = JSON.parse($("#tmaSelector").val());

        selectTMA(selectedTMA);
    }

    function selectTMA(selectedTMA) {
        var url = selectedTMA.Url;
        console.log(url);
        $("#etmaRequirements").attr("href",url);


    }


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

                // Clear out the previous header
                myHeader.clear();

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
        document.getElementById('message').innerText = message;
    }

    function getHeaderText() {
         return userDetails.Name + " (PI " + userDetails.PI + ")   "+ currentCourseCode +" - " + currentTMA;
    }

    function submitETMA() {

        hide($('#tools'));
        show($('#submitETMAPanel'));
        setSubmitButtonEnabled();
    }

    function setSubmitButtonEnabled() {
        $('#okSubmit').attr("disabled",! $('#chk1').is(':checked') && $('#chkCorrectFormat').is(':checked') && $('#chkNoCopying').is(':checked'));
    }

    function sendSubmission() {

    }



    function cancelSubmission() {
        show($('#tools'));
        hide($('#submitETMAPanel'));

    }
})();