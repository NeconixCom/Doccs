/// <reference path="/Scripts/FabricUI/MessageBanner.js" />
/// <reference path="/Scripts/marked.js" />


(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            initialiseUserInterface();

            var content = Office.context.document.settings.get("com.neconix.doccs.content");

            if (!content || content === "") {
                navigateToWelcomePage();
            } else {
                var renderedContent = marked(content);
                $("#content-text").html(renderedContent);
                navigateToReadingPage();
            }
        });
    };

    // User interface

    function initialiseUserInterface() {
        $("#content-placeholder-button").click(onContentPlaceholderButtonClicked);
        $("#toolbar-edit-button").click(onToolbarEditButtonClicked);
        $("#toolbar-done-button").click(onToolbarDoneButtonClicked);
        $("#toolbar-cancel-button").click(onToolbarCancelButtonClicked);
    }

    function onContentPlaceholderButtonClicked() {
        var content = Office.context.document.settings.get("com.neconix.doccs.content");

        if (!content) {
            content = "";
        }
        
        $("#content-editable-text-field").val(content);

        navigateToEditingPage();

        $("#content-editable-text-field").focus();
    }

    function onToolbarEditButtonClicked() {
        var content = Office.context.document.settings.get("com.neconix.doccs.content");

        if (!content) {
            content = "";
        }

        $("#content-editable-text-field").val(content);

        navigateToEditingPage();

        $("#content-editable-text-field").focus();
    }

    function onToolbarDoneButtonClicked() {
        var content = $("#content-editable-text-field").val();
        Office.context.document.settings.set("com.neconix.doccs.content", content);
        Office.context.document.settings.saveAsync();

        if (!content || content === "") {
            navigateToWelcomePage();
        } else {
            var renderedContent = marked(content);
            $("#content-text").html(renderedContent);
            navigateToReadingPage();
        }
    }

    function onToolbarCancelButtonClicked() {
        var content = Office.context.document.settings.get("com.neconix.doccs.content");

        if (!content || content === "") {
            navigateToWelcomePage();
        } else {
            var renderedContent = marked(content);
            $("#content-text").html(renderedContent);
            navigateToReadingPage();
        }
    }

    // Page navigation

    function navigateToWelcomePage() {
        showWelcomePage();
        hideReadingPage();
        hideEditingPage();
    }

    function navigateToReadingPage() {
        hideWelcomePage();
        showReadingPage();
        hideEditingPage();
    }

    function navigateToEditingPage() {
        hideWelcomePage();
        hideReadingPage();
        showEditingPage();
    }

    // Page visibility

    function showWelcomePage() {
        $("#page-welcome").show();
    }

    function hideWelcomePage() {
        $("#page-welcome").hide();
    }

    function showReadingPage() {
        $("#page-reading").show();
    }

    function hideReadingPage() {
        $("#page-reading").hide();
    }

    function showEditingPage() {
        $("#page-editing").show();
    }

    function hideEditingPage() {
        $("#page-editing").hide();
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
