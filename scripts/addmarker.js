/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
function getToolbarItems() {
    return {
        guid: window.Asc.plugin.info.guid,
        tabs: [{
            id: "tab_1",
            lockInViewMode: true,
            text: "Разметка документа",
            items: [
                {
                    id: "insertStamp",
                    type: "big-button",
                    text: "Место штампа",
                    // hint: "insert text into the document",
                    icons: "resources/img/stamp.svg",
                    lockInViewMode: true,
                    enableToggle: false,
                    separator: false
                },
                {
                    id: "insertDirectorName",
                    type: "button",
                    text: "ФИО подписанта",
                    // hint: "Insert an OLE object into the document",
                    icons: "resources/img/user.svg",
                    lockInViewMode: true,
                    enableToggle: false,
                    separator: false
                },
                {
                    id: "insertDocumentName",
                    type: "button",
                    text: "Название документа",
                    // hint: "Insert an image into the document",
                    icons: "resources/img/document.svg",
                    lockInViewMode: true,
                    enableToggle: false,
                    separator: false
                }
            ]
        }]
    };
}


function insertStamp(pl) {
    pl.callCommand(function () {
        var oDocument = Api.GetDocument();

        // Create a new paragraph
        var oParagraph = Api.CreateParagraph();

        // Add text to the paragraph
        oParagraph.AddText("${image_stamp}");

        // Style the text as a title
        oParagraph.SetBold(true);              // Make the text bold
        // oParagraph.SetFontSize(24);            // Increase the font size
        // oParagraph.SetJc("center");         // Align text to the center

        // Insert the paragraph at the beginning of the document
        oDocument.InsertContent([oParagraph], 0);
    });
}

function insertDirectorName(pl) {
    pl.callCommand(function () {
        var oDocument = Api.GetDocument();

        // Create a new paragraph
        var oParagraph = Api.CreateParagraph();

        // Add text to the paragraph
        oParagraph.AddText("${director.name}");

        // Style the text as a title
        oParagraph.SetBold(true);              // Make the text bold
        // oParagraph.SetFontSize(24);            // Increase the font size
        // oParagraph.SetJc("center");         // Align text to the center

        // Insert the paragraph at the beginning of the document
        oDocument.InsertContent([oParagraph], 0);
    });
}



function insertDocumentName(pl) {
    pl.callCommand(function () {
        var oDocument = Api.GetDocument();

        // Create a new paragraph
        var oParagraph = Api.CreateParagraph();

        // Add text to the paragraph
        oParagraph.AddText("${document.name}");

        // Style the text as a title
        oParagraph.SetBold(true);              // Make the text bold
        // oParagraph.SetFontSize(24);            // Increase the font size
        // oParagraph.SetJc("center");         // Align text to the center

        // Insert the paragraph at the beginning of the document
        oDocument.InsertContent([oParagraph], 0);
    });
}
// Example insert text into editors (different implementations)
(function (window, undefined) {


    // window.Asc.plugin.attachEditorEvent("onToolbarMenuClick", (data) => {
    //     console.log("insertStamp");
    // });


    // window.Asc.plugin.attachToolbarMenuClickEvent("tab_1", (data) => {
    //     console.log("insertStamp");
    // });


    // window.Asc.plugin.attachToolbarMenuClickEvent("insertStamp", function (data) {
    //     console.log("insertStamp");
    // });

    // var text = "${image_stamp}";
    // window.Asc.plugin.event_onToolbarMenuClick = (id) => {
    //     console.log("Toolbar menu item clicked: " + id);
    // };

    window.Asc.plugin.init = function () {
        this.executeMethod("AddToolbarMenuItem", [getToolbarItems()])

        this.attachEditorEvent("onToolbarMenuClick", (data) => {
            switch (data) {
                case 'insertStamp':
                    insertStamp(this);
                    break;
                case 'insertDirectorName':
                    insertDirectorName(this);
                    break;
                case 'insertDocumentName':
                    insertDocumentName(this);
                    break;
            }
        });
    };

})(window, undefined);
