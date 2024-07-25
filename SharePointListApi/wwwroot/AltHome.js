'use strict';

// This code serves as an example to do more complex actions with the Office app and the disclaimer list, such as
// picking multiple responses and inserting them into the document as a header or footer.
// To experiment with this code, you need to map the manifest for Word to point to the AltHome.html file instead of Home.html.

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

            // Add event listener for the new button
            $('#insertSelectedDisclaimersAsHeader').on('click', function () {
                insertSelectedDisclaimersAsHeader();
            });

            $('#insertSelectedDisclaimersAsFooter').on('click', function () {
                insertSelectedDisclaimersAsFooter();
            });

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

                // Create a <ul> element to hold the list of disclaimers with checkboxes
                let list = $('<ul style="list-style-type: none;"></ul>');
                disclaimerList.append(list);


                disclaimers.forEach(disclaimer => {
                    // Create a list item for each disclaimer
                    let listItem = $('<li></li>');

                    // Create a checkbox and label for the disclaimer
                    let checkbox = $(`<input type="checkbox" id="disclaimer_${disclaimer.id}" value="${disclaimer.text}">`);
                    let label = $(`<label for="disclaimer_${disclaimer.id}">${disclaimer.description}</label>`);

                    // Append the checkbox and label to the list item
                    listItem.append(checkbox);
                    listItem.append(label);

                    // Append the list item to the list
                    list.append(listItem);
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

                // Create a <ul> element to hold the list of disclaimers with checkboxes
                let list = $('<ul style="list-style-type: none;"></ul>');
                disclaimerList.append(list);

                disclaimers.forEach(disclaimer => {
                    // Create a list item for each disclaimer
                    let listItem = $('<li></li>');

                    // Create a checkbox and label for the disclaimer
                    let checkbox = $(`<input type="checkbox" id="disclaimer_${disclaimer.id}" value="${disclaimer.text}">`);
                    let label = $(`<label for="disclaimer_${disclaimer.id}">${disclaimer.description}</label>`);

                    // Append the checkbox and label to the list item
                    listItem.append(checkbox);
                    listItem.append(label);

                    // Append the list item to the list
                    list.append(listItem);
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

    // A simple way to determine the Office app context of the running Add-in, you can check the value of Office.context.host.
    // Office.HostType.Outlook is the host type for Outlook.
    // Office.HostType.Word is the host type for Word.
    // Office.HostType.Excel is the host type for Excel.
    // Office.HostType.PowerPoint is the host type for PowerPoint.

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

    function insertSelectedDisclaimersAsHeader() {
        let selectedDisclaimersText = '';
        $('#disclaimer-list input[type="checkbox"]:checked').each(function () {
            // Assuming the value of each checkbox is the text you want to insert
            selectedDisclaimersText += $(this).val() + '\n'; // Collecting text and adding a newline after each
        });

        if (selectedDisclaimersText) {

            // Insert the collected text into the header of the document
            Word.run(function (context) {
                // Get the primary header of the first section.
                var header = context.document.sections.getFirst().getHeader("primary");

                // Insert text into the header.
                header.insertText(selectedDisclaimersText, "Replace");

                return context.sync()
                    .then(function () {
                        console.log('Header added to the first section.');
                    });
            }).catch(function (error) {
                console.log('Error:', error);
            });
        } else {
            console.log('No disclaimers selected.');
        }
    }

    function insertSelectedDisclaimersAsFooter() {
        let selectedDisclaimersText = '';
        $('#disclaimer-list input[type="checkbox"]:checked').each(function () {
            // Assuming the value of each checkbox is the text you want to insert
            selectedDisclaimersText += $(this).val() + '\n'; // Collecting text and adding a newline after each
        });

        if (selectedDisclaimersText) {
            // Insert the collected text into the footer of the document
            Word.run(function (context) {
                // Get the primary footer of the first section.
                var footer = context.document.sections.getFirst().getFooter("primary");

                // Insert text into the footer.
                footer.insertText(selectedDisclaimersText, "Replace");

                return context.sync()
                    .then(function () {
                        console.log('Footer added to the first section.');
                    });
            }).catch(function (error) {
                console.log('Error:', error);
            });
        } else {
            console.log('No disclaimers selected.');
        }

    }


})();
