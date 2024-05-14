var agileHeaders = null;
var custHeaders = null;

function addMappingOption(optionFor, idPrefix, headersDict) {
    const idHeaders = ["level", "part-number", "description", "uom", "qty", "rev", "ref-des", "mfr-name", "mfr-number"];
    idHeaders.forEach(function (idHeader) {
        var selectElement = null;
        if (optionFor == "cust" && (idHeader == "mfr-name" || idHeader == "mfr-number")) {
            selectElement = document.getElementById(idPrefix + "-" + idHeader + "-1");
        } else {
            selectElement = document.getElementById(idPrefix + "-" + idHeader);
        }
        selectElement.disabled = false;
        Object.keys(headersDict).forEach(function (header) {
            let columnIndex = headersDict[header];
            let option = document.createElement("option");
            option.setAttribute("value", columnIndex);
            option.innerText = header;
            selectElement.appendChild(option);
        });
    });
}

function addSupplierOption(headersDict, selectTag) {
    // Add default/None option
    var defaultOption = document.createElement("option");
    defaultOption.setAttribute("value", "-1");
    defaultOption.innerText = "None";
    defaultOption.selected = true;
    selectTag.appendChild(defaultOption);
    Object.keys(headersDict).forEach(function (header) {
        let columnIndex = headersDict[header];
        let option = document.createElement("option");
        option.setAttribute("value", columnIndex);
        option.innerText = header;
        selectTag.appendChild(option);
    });
}

function addSupplierSelectTag(parentDOM, selectTagId, selectTagIndex, textLabel, rowId) {
    // create select tags
    const selectTag = document.createElement("select");
    selectTag.setAttribute("id", selectTagId + selectTagIndex);
    selectTag.setAttribute("name", selectTagId + selectTagIndex);
    selectTag.required = true;
    const tdSelect = document.createElement("td");
    addSupplierOption(custHeaders, selectTag);
    tdSelect.appendChild(selectTag);

    // create Mfr Name label
    const labelTag = document.createElement("label");
    labelTag.innerText = textLabel + " " + selectTagIndex;
    const tdLabel = document.createElement("td");
    tdLabel.setAttribute("for", selectTagId + selectTagIndex);
    tdLabel.appendChild(labelTag);

    // create Mfr Name row
    const trTag = document.createElement("tr");
    // Note: row id will be used to remove the row if number of suppliers decrease
    trTag.setAttribute("id", rowId + "-" + selectTagIndex);
    trTag.appendChild(tdLabel);
    trTag.appendChild(tdSelect);

    parentDOM.append(trTag);
}

$(document).ready(function () {
    $("#upload-template-field").change(function (event) {
        var form_data = new FormData(); // Get file data from upload field
        form_data.append("file", $("#upload-template-field").prop("files")[0]);

        $.ajax({
            type: "POST",
            url: "/load_raw_bom_template",
            data: form_data,
            contentType: false,
            cache: false,
            processData: false,
            success: function (response) {
                agileHeaders = response[0];
                custHeaders = response[1];

                // Generate Agile Mapping Section
                // var $agileMappingTBody = $("#agile-bom-mapping-table tbody");
                addMappingOption("agile", "agile-bom-mapping", agileHeaders);
                // Genrate Customer Mapping Section
                addMappingOption("cust", "cust-bom-mapping", custHeaders);
            },
        }); // end $.ajax()
    });

    // Handle the change in number-of-supplier input
    // and add more supplier box to the form
    (function () {
        var $numberOfSupplierInput = $("#number-of-supplier");
        $numberOfSupplierInput.change(function () {
            var totalSupplierDataSet = document.getElementById("total-suppliers");
            var totalSuppliers = parseInt(totalSupplierDataSet.dataset.suppliers);

            var numberOfSupplier = parseInt($numberOfSupplierInput.val());
            var $custBomMappingTable = $("#customer-bom-mapping-table tbody");

            const mfrNameSelectId = "cust-bom-mapping-mfr-name-";
            const mfrNumberSelectId = "cust-bom-mapping-mfr-number-";

            if (numberOfSupplier > totalSuppliers) {
                for (var i = totalSuppliers; i < numberOfSupplier; i++) {
                    var index = i + 1;
                    // Add Mfr Name selection
                    addSupplierSelectTag($custBomMappingTable, mfrNameSelectId, index, "Mfr Name", "mfr-name");
                    // Add Mfr Number
                    addSupplierSelectTag($custBomMappingTable, mfrNumberSelectId, index, "Mfr Number", "mfr-number");
                }
            } else if (numberOfSupplier < totalSuppliers) {
                for (var i = totalSuppliers; i > numberOfSupplier; i--) {
                    document.getElementById("mfr-number-" + i).remove();
                    document.getElementById("mfr-name-" + i).remove();
                }
            }
            document.querySelector("#total-suppliers").dataset['suppliers'] = numberOfSupplier;
            document.getElementById("total-suppliers").setAttribute("value", numberOfSupplier);
        });
    }());
});
