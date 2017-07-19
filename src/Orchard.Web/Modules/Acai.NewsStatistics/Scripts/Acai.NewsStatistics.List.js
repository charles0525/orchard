var exportExcel={
    init:function(){
        $(function(){
            $('.export-btn').on('click',exportExcel.createExportFile);            
        });
    },
    createExportFile:function(layerId){
        if(!layerId){
            layerId='news-layer';
        }

        var params ={__RequestVerificationToken:$('[name="__RequestVerificationToken"]').val()};
        $.ajax({
            type: "post",
            dataType: "json",
            data: params,
            url: '/Acai.NewsStatistics/Admin/ExportPost',
            success: function (data) {
                var json = data;
                if (json != null && json.status==1) {
                    if (!!json.msg) {
                        exportExcel.downloadFile(layerId, json.msg);
                    }
                }
                else {
                    alert(json.msg);
                }
            },
            beforeSend: function (XHR) {
            },
            complete: function (XHR, TS) {
            }
        });
    },
    downloadFile:function(layerId, file){
        var iframe = document.getElementById(layerId + "_iframe");
        if (iframe != undefined && iframe != null) {
            document.body.removeChild(iframe);
        }

        iframe = document.createElement("iframe");
        iframe.id = layerId + "_iframe";
        iframe.src = file;
        iframe.style.display = "none";
        document.body.appendChild(iframe);
    }
}
exportExcel.init();