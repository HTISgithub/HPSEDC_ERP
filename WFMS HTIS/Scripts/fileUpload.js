$(function(){
	$('body').on('submit','#frm',function(e){
		e.preventDefault();
		
		var url = $(this).attr('action');
		var frm = $(this);
		
		var data = new FormData();
		if(frm.find('#txtFile[type="file"]').length === 1 ){
		    data.append('file', frm.find('#postedfile')[0].files[0]);
		}
		
		 
		var ajax  = new XMLHttpRequest();
		ajax.upload.addEventListener('progress',function(evt){
			var percentage = (evt.loaded/evt.total)*100;
			upadte_progressbar(Math.round(percentage));
		},false);
		ajax.addEventListener('load',function(evt){
			if(evt.target.responseText.toLowerCase().indexOf('error')>=0){
				show_error(evt.target.responseText);
			}else{
				preview_image(evt.target.responseText);
			}
			upadte_progressbar(0);
			frm[0].reset();
			
		},false);
		ajax.addEventListener('error',function(evt){
			show_error('upload failed');
			upadte_progressbar(0);
		},false);
		ajax.addEventListener('abort',function(evt){
			show_error('upload aborted');
			upadte_progressbar(0);
		},false);
		ajax.open('POST',url);
		ajax.send(data);
		return false;
	}); 
});

function upadte_progressbar(value){
	$('#progressBar').css('width',value+'%').html(value+'%');
	if(value==0){
		$('.progress').hide();
	}else{
		$('.progress').show();
	}
}

function preview_image(src){
	$('.thumbnail').show();
	var d = new Date();
	$('#img-preview').removeAttr('src').attr('alt','loading image');
	$('#img-preview').attr('src',src+'?random='+d.getTime());
}

function show_error(error){
	$('.thumbnail, #progressBar').hide();
	$('#error').show();
	$('#error').html(error);
}