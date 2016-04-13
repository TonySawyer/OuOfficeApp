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
      //  Office.context.document.

    }
})();