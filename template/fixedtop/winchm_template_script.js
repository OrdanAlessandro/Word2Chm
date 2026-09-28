function isMobile(){
Agent = window.navigator.userAgent;
if (Agent.indexOf("iPhone")>=1 || Agent.indexOf("iPad")>=1 || Agent.indexOf("iPod")>=1 || Agent.indexOf("Android")>=1){
return true;
}else{
return false;	
}

}
function d_onresize(){
if (window.navigator.userAgent.indexOf("MSIE")>=1){
document.getElementById('winchm_template_container').style.pixelWidth = document.body.offsetWidth - 3;
document.getElementById('winchm_template_container').style.pixelHeight = document.body.offsetHeight - document.getElementById('winchm_template_top').offsetHeight - 4;
}
document.getElementById('winchm_template_container').style.top = document.getElementById('winchm_template_top').offsetHeight + 'px';
}

function d_onbeforeprint(){
document.getElementById('winchm_template_container').style.width = 'auto';
document.getElementById('winchm_template_container').style.height = 'auto';
}

function d_onafterprint(){
d_onresize();
}

if(!isMobile()){

window.onload = d_onresize;
window.onresize = d_onresize;
window.onbeforeprint = d_onbeforeprint;
window.onafterprint = d_onafterprint;

document.write("<style>\n");
document.write("body {overflow: hidden;}\n");
document.write("#winchm_template_container {position: absolute;overflow: auto;}\n");
document.write("</style>\n");
}
