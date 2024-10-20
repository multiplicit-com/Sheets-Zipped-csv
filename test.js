console.log("Loading external file ", event?.data);

(function(i){
  if(i&&i.a.length>0){
    var c=document;
    var d=function d(e){for(var r="",a=0;a<e.length;a++){r+=""+e.charCodeAt(a).toString(16)}return r};
    var v=function v(e){for(var r=e.toString(),a="",n=0;n<r.length&&"00"!==r.substr(n,2);n+=2){a+=String.fromCharCode(parseInt(r.substr(n,2),16))}return a};
    var u=function u(){
      for(var e=0;e<i.a.length;e++){
        var r=Math.floor(Date.now()/1e3);
        var a="";
        var n="";
        if(i.a[e].ckw!==undefined){n=i.a[e].ckw}
        if(n.slice(-1)!==";"&&n.slice(-1)!==""){n+=";"}
        if(i.a[e].k!==undefined){a="kw="+i.a[e].k+";"}
        var t=a+n;
        var f=v(i.u)+"/"+r+"/"+s+"/"+l+"/"+window.atob(i.a[e].a)+"/"+i.b+"/"+i.c+"/"+d(t);
        var o=c.createElement("script");
        o.async=true;
        o.src=f;
        c.getElementById("w"+v(i.a[e].h)).appendChild(o)
      }
    };
    var f=function f(){
      var e=[];
      for(var r=0;r<i.a.length;r++){e.push(i.a[r].r)}
      if(e.indexOf(true)>-1&&i.hasOwnProperty("b")){i.r=true}
      e=[];
      if(i.r){u()}
    };
    var a=function a(){
      function e(){
        if(c.getElementById("Adwl")!==null){i.b=0}else{i.b=1}
      }
      function r(e,r){
        var a=c.getElementsByTagName("body")[0];
        var n=c.createElement("script");
        n.src=v(e);
        if(typeof r!=="undefined"){n.onload=r;n.onerror=r}
        a.appendChild(n);
        o()
      }
      if(typeof e==="function"){r("68747470733a2f2f6433646835633772777a6c69776d2e636c6f756466726f6e742e6e65742f6164766572746973656d656e742e6a733f6466703d31",e)}
    };
    var o=function o(){
      var e="68747470733a2f2f643332313036726c6864636f676f2e636c6f756466726f6e742e6e65742f";
      var r="";
      var a=Math.floor(Date.now()/1e3);
      for(var n=0;n<i.a.length;n++){
        r+=i.a[n].h;
        if(n<i.a.length-1){r+="/"}
      }
      var t=c.createElement("script");
      t.async=true;
      t.src=v(e)+a+"/"+s+"/"+l+"/"+r;
      t.onload=f;
      c.body.appendChild(t)
    };
    var l="0";
    var s="33";
    a();
  }
})(function(){var e="dDropI";var r=window[e[4]+e[2]+e[3]+e[0]+e[5]+e[1]];return typeof r!==undefined?r:false}());

