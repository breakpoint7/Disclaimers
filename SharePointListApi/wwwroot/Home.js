'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready

            // The code below adds two buttons that load a list of responses from different sources, one is a Web API and another is a SharePoint list.
            // Depending on the way you choose to implement this, you might remove the option and immediately load responses from one or the other (removing the option to choose).)
            let disclaimerSources = $('#disclaimer-source');

            // Add the option to load responses from a Web API (hard coded in the sample)
            let disclaimerSourcesElementOData = $(`<button class="ms-Button">Web API</button>`);
            disclaimerSourcesElementOData.on("click", function () {
                fetchDisclaimers();
            });
            disclaimerSources.append(disclaimerSourcesElementOData); // Fixed variable name and moved inside the function

            // Add the option to load response from a SharePoint list
            let disclaimerSourcesElementSharePoint = $(`<button class="ms-Button">SharePoint List</button>`);
            disclaimerSourcesElementSharePoint.on("click", function () {
                fetchDisclaimersSharePointList();
            });
            disclaimerSources.append(disclaimerSourcesElementSharePoint); // Fixed variable name and moved inside the function
        });
    });

    // Fetch disclaimers from a Web API
    function fetchDisclaimers() {
        fetch('https://localhost:7057/api/disclaimer')
            .then(response => response.json())
            .then(data => {
                let disclaimers = data.value;
                let disclaimerList = $('#disclaimer-list');
                disclaimerList.empty();
                disclaimers.forEach(disclaimer => {
                    let disclaimerElementWrapper = $('<div class="disclaimer-button-wrapper"></div>');
                    let disclaimerElement = $(`<button class="ms-Button">${disclaimer.description}</button>`);
                    disclaimerElement.on("click", function () {
                        insertText(disclaimer.text);
                    });
                    disclaimerElementWrapper.append(disclaimerElement);
                    disclaimerList.append(disclaimerElementWrapper);
                });
            })
            .catch(error => {
                console.log('Error:', error);
                showMessage(`Failed to load disclaimers from Web API: ${error.message}`);
            });
    }

    // Fetch disclaimers from a SharePoint list
    // This calls a web API that acts as a intermediary to fetch data from SharePoint (so we can add app level credentials to authorize the request)
    async function fetchDisclaimersSharePointList() {

        fetch('https://localhost:7057/api/sharepointlist')
            .then(response => response.json())
            .then(data => {
                let disclaimers = data.value;
                let disclaimerList = $('#disclaimer-list');
                disclaimerList.empty();
                disclaimers.forEach(disclaimer => {
                    let disclaimerElementWrapper = $('<div class="disclaimer-button-wrapper"></div>');
                    let disclaimerElement = $(`<button class="ms-Button">${disclaimer.description}</button>`);
                    disclaimerElement.on("click", function () {
                        insertText(disclaimer.text);
                    });
                    disclaimerElementWrapper.append(disclaimerElement);
                    disclaimerList.append(disclaimerElementWrapper);
                });
            })
            .catch(error => {
                console.log('Error:', error);
                showMessage(`Failed to load disclaimers from SharePoint List: ${error.message}`);
            });

    }


    function showMessage(text) {
        const appendedText = $('#disclaimer-list').html() + text + "<br>---";
        $('#disclaimer-list').html(appendedText); // Targeting #disclaimer-list for displaying the message
    }

    function insertText(text) {
        // Check if the host is Outlook
        if (Office.context.mailbox) {
            // Use the setSelectedDataAsync method on the body object to insert text
            Office.context.mailbox.item.body.setSelectedDataAsync(text,
                { coercionType: Office.CoercionType.Text },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                }
            );
        } else {
            // Fallback for other Office applications
            Office.context.document.setSelectedDataAsync(text,
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                }
            );
        }
    }


})();
