function onOpen() {
    var weekly = new LoadWeekly(2);
    weekly.run(loadMonthly);

    function loadMonthly(_resultCurrentRowId) {
        var monthly = new LoadMonthly(_resultCurrentRowId);
        monthly.run();
    }
}