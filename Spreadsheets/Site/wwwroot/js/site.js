$(document)
    .ready(function(e) {
        $(".current-fees-refresh")
            .click(function(e) {
                updateTable();
            });

        function updateTable() {
            $.ajax("Home/CurrentFees").done(function(data) {
                $("#spreadsheet").html(data);
            });
        }
    });