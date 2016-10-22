﻿/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(
                    displaySelectedText);
                return;
            }

            $("#template-description").text("欢迎参加微软 DevDays 2016 活动！ 本插件用于演示 Office Add-in 的一些主要功能，比如 Dialog API， Word API， 等等。");
            $('#button-text').text("高亮");
            $('#button-desc').text("高亮显示所选范围的最长单词");
            $('#openDialog-text').text("弹出窗口");
            $('#openDialog-desc').text("弹出新的对话框窗口");
            $('#buttonInsertPicture-text').text("插图");
            $('#buttonInsertPicture-desc').text("将选择的照片插入到文档指定位置");
            
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $('#openDialog-button').click(openDialog);
        });
    };

    function loadSampleData() {

        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText("This is a sample text inserted in the document",
                            Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightLongestWord() {

        Word.run(function (context) {

            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // variable for keeping the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = context.document.body.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');

                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync)
        })
        .catch(errorHandler);
    } 

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
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

    function openDialog() {
        //var url = "https://www.bing.com";
        var url = "https://beetroot.azurewebsites.net/Dialog.html"

        Office.context.ui.displayDialogAsync(url, { height: 60, width: 50, displayInIframe: true }, dialogCallback);
    }

