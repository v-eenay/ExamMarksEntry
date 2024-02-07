const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
$(document).ready(function () {
  $("#excelProcessorForm").submit(function (e) {
    e.preventDefault(); // Prevent form submission

    // Get form data
    var courseworkType = $("#courseworkType").val();
    var folderPath = $("#folderPath").val();
    var individualSheetName = $("#individualSheetName").val();
    var destinationSheetName = $("#destinationSheetName").val();
    var groupNameColumn = $("#groupNameColumn").val();
    var studentIDCell = $("#studentIDCell").val();
    var studentIDColumn = $("#studentIDColumn").val();
    var individualMarksCells = $("input[name='individualMarksCell[]']")
      .map(function () {
        return $(this).val();
      })
      .get();
    var destinationColumns = $("input[name='destinationColumn[]']")
      .map(function () {
        return $(this).val();
      })
      .get();
    var destinationSheetFile = $("#destinationSheet").prop("files")[0];

    // Validation
    if (
      !folderPath ||
      !individualSheetName ||
      !destinationSheetName ||
      !studentIDCell ||
      !studentIDColumn ||
      individualMarksCells.length !== destinationColumns.length
    ) {
      alert(
        "Please fill in all the required fields and ensure each individual marks cell has a corresponding destination column."
      );
      return;
    }

    // Logic to process marks based on coursework type
    if (courseworkType === "individual") {
      processIndividualCoursework(
        folderPath,
        individualSheetName,
        destinationSheetName,
        studentIDCell,
        studentIDColumn,
        individualMarksCells,
        destinationColumns,
        destinationSheetFile
      );
    } else if (courseworkType === "group") {
      processGroupCoursework(
        folderPath,
        individualSheetName,
        destinationSheetName,
        groupNameColumn,
        studentIDCell,
        studentIDColumn,
        individualMarksCells,
        destinationColumns,
        destinationSheetFile
      );
    }
  });
});

function processIndividualCoursework(
  folderPath,
  individualSheetName,
  destinationSheetName,
  studentIDCell,
  studentIDColumn,
  individualMarksCells,
  destinationColumns,
  destinationSheetFile
) {
  // Load destination workbook
  var reader = new FileReader();
  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: "array" });

    // Load destination sheet
    var destinationSheet = workbook.Sheets[destinationSheetName];
    if (!destinationSheet) {
      alert(
        "Destination sheet '" +
          destinationSheetName +
          "' not found in the destination file."
      );
      return;
    }

    // Read and parse student IDs and corresponding row numbers
    var studentIDRows = {};
    var studentIDColumnIndex = XLSX.utils.decode_col(studentIDColumn);
    var studentIDCellAddress = XLSX.utils.decode_cell(studentIDCell);
    var studentIDRow = studentIDCellAddress.r;
    var studentIDCol = studentIDCellAddress.c;
    while (
      destinationSheet[
        XLSX.utils.encode_cell({ r: studentIDRow, c: studentIDColumnIndex })
      ]
    ) {
      var studentID =
        destinationSheet[
          XLSX.utils.encode_cell({ r: studentIDRow, c: studentIDColumnIndex })
        ].v;
      studentIDRows[studentID] = studentIDRow;
      studentIDRow++;
    }

    // Load individual files from folder path
    $.ajax({
      url: folderPath,
      success: function (data) {
        $(data)
          .find("a")
          .attr("href", function (i, filename) {
            if (filename.endsWith(".xlsx") || filename.endsWith(".xls")) {
              var fileURL = folderPath + "/" + filename;
              var reader = new FileReader();
              reader.onload = function (e) {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: "array" });
                var sheet = workbook.Sheets[individualSheetName];
                if (!sheet) {
                  alert(
                    "Individual sheet '" +
                      individualSheetName +
                      "' not found in file '" +
                      filename +
                      "'."
                  );
                  return;
                }

                // Read student ID and marks from individual sheet
                var studentID =
                  sheet[XLSX.utils.encode_cell(studentIDCellAddress)].v;
                var marks = {};
                for (var i = 0; i < individualMarksCells.length; i++) {
                  var cellAddress = XLSX.utils.decode_cell(
                    individualMarksCells[i]
                  );
                  marks[destinationColumns[i]] =
                    sheet[XLSX.utils.encode_cell(cellAddress)].v;
                }

                // Write marks to destination sheet
                var destinationSheetRow = studentIDRows[studentID];
                if (!destinationSheetRow) {
                  alert(
                    "Student ID '" +
                      studentID +
                      "' not found in destination sheet."
                  );
                  return;
                }
                for (var destColumn in marks) {
                  var destColumnIndex = XLSX.utils.decode_col(destColumn);
                  destinationSheet[
                    XLSX.utils.encode_cell({
                      r: destinationSheetRow,
                      c: destColumnIndex,
                    })
                  ] = { v: marks[destColumn] };
                }

                // Write modified destination sheet to destination file
                var wbout = XLSX.write(workbook, {
                  bookType: "xlsx",
                  type: "binary",
                });
                saveAs(
                  new Blob([s2ab(wbout)], { type: "application/octet-stream" }),
                  "Processed_" + filename
                );
              };
              reader.readAsArrayBuffer(fileURL);
            }
          });
      },
      error: function (xhr, textStatus, errorThrown) {
        alert("Error loading individual files from folder: " + errorThrown);
      },
    });
  };
  reader.readAsArrayBuffer(destinationSheetFile);
}

function processGroupCoursework(
  folderPath,
  individualSheetName,
  destinationSheetName,
  groupNameColumn,
  studentIDCell,
  studentIDColumn,
  individualMarksCells,
  destinationColumns,
  destinationSheetFile
) {
  var reader = new FileReader();
  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: "array" });

    // Load destination sheet
    var destinationSheet = workbook.Sheets[destinationSheetName];
    if (!destinationSheet) {
      alert(
        "Destination sheet '" +
          destinationSheetName +
          "' not found in the destination file."
      );
      return;
    }

    // Read and parse student IDs and corresponding row numbers
    var studentIDRows = {};
    var studentIDColumnIndex = XLSX.utils.decode_col(studentIDColumn);
    var studentIDCellAddress = XLSX.utils.decode_cell(studentIDCell);
    var studentIDRow = studentIDCellAddress.r;
    while (
      destinationSheet[
        XLSX.utils.encode_cell({ r: studentIDRow, c: studentIDColumnIndex })
      ]
    ) {
      var studentID =
        destinationSheet[
          XLSX.utils.encode_cell({ r: studentIDRow, c: studentIDColumnIndex })
        ].v;
      studentIDRows[studentID] = studentIDRow;
      studentIDRow++;
    }

    // Load group folders from folder path
    $.ajax({
      url: folderPath,
      success: function (data) {
        $(data)
          .find("a")
          .attr("href", function (i, groupName) {
            if (groupName !== "." && groupName !== "..") {
              var groupFolderPath = folderPath + "/" + groupName;
              var groupReader = new FileReader();
              groupReader.onload = function (groupEvent) {
                var groupData = new Uint8Array(groupEvent.target.result);
                var groupWorkbook = XLSX.read(groupData, { type: "array" });

                // Load individual sheet from group workbook
                var groupSheet = groupWorkbook.Sheets[individualSheetName];
                if (!groupSheet) {
                  alert(
                    "Individual sheet '" +
                      individualSheetName +
                      "' not found in group folder '" +
                      groupName +
                      "'."
                  );
                  return;
                }

                // Read student ID and marks from individual sheet
                var studentID =
                  groupSheet[XLSX.utils.encode_cell(studentIDCellAddress)].v;
                var marks = {};
                for (var i = 0; i < individualMarksCells.length; i++) {
                  var cellAddress = XLSX.utils.decode_cell(
                    individualMarksCells[i]
                  );
                  marks[destinationColumns[i]] =
                    groupSheet[XLSX.utils.encode_cell(cellAddress)].v;
                }

                // Write marks to destination sheet
                var destinationSheetRow = studentIDRows[studentID];
                if (!destinationSheetRow) {
                  alert(
                    "Student ID '" +
                      studentID +
                      "' not found in destination sheet."
                  );
                  return;
                }
                for (var destColumn in marks) {
                  var destColumnIndex = XLSX.utils.decode_col(destColumn);
                  destinationSheet[
                    XLSX.utils.encode_cell({
                      r: destinationSheetRow,
                      c: destColumnIndex,
                    })
                  ] = { v: marks[destColumn] };
                }

                var wbout = XLSX.write(workbook, {
                  bookType: "xlsx",
                  type: "binary",
                });
                saveAs(
                  new Blob([s2ab(wbout)], { type: "application/octet-stream" }),
                  "Processed_" + groupName + ".xlsx"
                );
              };
              groupReader.readAsArrayBuffer(groupFolderPath);
            }
          });
      },
      error: function (xhr, textStatus, errorThrown) {
        alert("Error loading group folders from folder: " + errorThrown);
      },
    });
  };
  reader.readAsArrayBuffer(destinationSheetFile);
}
