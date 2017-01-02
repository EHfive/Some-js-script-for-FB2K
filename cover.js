//一个简单的获取网易云音乐封面的WSH脚本
//新建WSH面板载入脚本使用
//获取的封面保存在 歌曲目录\Cover[-%artist%]-%title%.jpg 下
//在WSH面板上显示获取的封面，可在参数选项：显示 搜索模板 在“cover.jpg”下添加“Cover[-%artist%]-%title%.jpg”
//以在封面显示器面板显示
//载入错误请重新加载
//mail:cimoc@sokka.cn

var xmlHttp = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
var debug=false;
var ww = 0,
    wh = 0;
var g_img = null;
var g_valid_tid = 0;
on_playlist_items_selection_change();
function on_playlist_items_selection_change(){

var song=fb.GetFocusItem(true);
var til = fb.TitleFormat("%title%").EvalWithMetadb(song).replace(/(\\|:|\*|\?|"|<|>|\/|\|)/g, "");
var art = fb.TitleFormat("%artist%").EvalWithMetadb(song).replace(/(\\|:|\*|\?|"|<|>|\/|\|)/g, "");
console(art);
var songpath=utils.FileTest(song.Path, "split").toArray();
var pathc=songpath[0]+"cover.jpg";
var tt=art?"Cover-"+art+"-"+til+".jpg":"Cover-"+til+".jpg";
var pathn = songpath[0]+tt;
if(utils.FileTest(pathn,"e")) var path=pathn;
else {
    if(utils.FileTest(pathc,"e")) var path=pathc;
    else {
           var path =pathn;
            var link=search(til,art);
    debug && console(link);
    DownURL(link, path);
        }
    }
load_image_async(path);
}

function load_image_async(path) {
    g_valid_tid = gdi.LoadImageAsync(window.ID, path);
}

function on_size() {
    ww = window.Width;
    wh = window.Height;
}

function on_paint(gr) {
    //gr.SetSmoothingMode(0);
    if (g_img) {
        var x = 0,
            y = 0;
        var scale = 0;
        var scale_w = ww / g_img.Width;
        var scale_h = wh / g_img.Height;

        if (scale_w <= scale_h) {
            scale = scale_w;
            y = (wh - g_img.Height * scale) / 2;
        } else {
            scale = scale_h;
            x = (ww - g_img.Width * scale) / 2;
        }

        gr.DrawImage(g_img, x, y, g_img.Width * scale, g_img.Height * scale, 0, 0, g_img.Width, g_img.Height);
    }
}

// After loading image is done in the background, this callback will be invoked
function on_load_image_done(tid, image) {
    if (g_valid_tid == tid) {
        // Dispose previous image, in order to save memory
        g_img && g_img.Dispose();
        g_img = image;
        window.Repaint();
    }
}
function DownURL(strRemoteURL, strLocalURL){
        try{
            var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
            xmlHTTP.open("Get", strRemoteURL, false);
            xmlHTTP.send();
            var adodbStream = new ActiveXObject("ADODB.Stream");
            adodbStream.Type = 1;//1=adTypeBinary 
            adodbStream.Open();
            adodbStream.write(xmlHTTP.responseBody);
            adodbStream.SaveToFile(strLocalURL, 2);
            adodbStream.Close();
            adodbStream = null;
            xmlHTTP = null;
        }
        catch (e){
            fb.trace(e);
        }
        
    }
function search(title,artist){
    var xmlHttp = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
    var searchURL, infoURL;
    var limit = 3;
    //删除feat.及之后内容并保存
    var str1 = del(title, "feat.");
    var str2 = del(artist, "feat.");
    var title = str1[0];
    var outstr1 = str1[1];
    var artist = str2[0];
    var outstr2 = str2[1];
    //搜索
    var s = artist ? (title + "-" + artist) : title;
    searchURL = "http://music.163.com/api/search/get/web?csrf_token=";//如果下面的没用,试试改成这句
    var post_data = 'hlpretag=<span class="s-fc7">&hlposttag=</span>&s=' + encodeURIComponent(s) + '&type=1&offset=0&total=true&limit=' + limit;
    try {
        xmlHttp.Open("POST", searchURL, false);
        //noinspection JSAnnotator
        xmlHttp.Option(4) = 13056;
        //noinspection JSAnnotator
        xmlHttp.Option(6) = false;
        xmlHttp.SetRequestHeader("Host", "music.163.com");
        xmlHttp.SetRequestHeader("Origin", "http://music.163.com");
        xmlHttp.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36");
        xmlHttp.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        xmlHttp.SetRequestHeader("Referer", "http://music.163.com/search/");
        xmlHttp.SetRequestHeader("Connection", "Close");
        xmlHttp.Send(post_data);
    } catch (e) {
        debug && console("Search failed");
    }
    if (xmlHttp.Status == 200) {
        //  console(xmlHttp.responseText);
        var ncm_back = json(xmlHttp.responseText);
        var result = ncm_back.result;
        if (ncm_back.code != 200 || !result.songCount) {
            debug && console("get info failed");
            return false;
        }
        //筛选曲名及艺术家
        var song = result.songs;
        var out = [0, 0];
        var b = 0;
        var c = 0;
        for (var k in song) {
            var ncm_name = song[k].name;
            for (var a_k in song[k].artists) {
                var ncm_artist = song[k].artists[a_k].name;
                var p0 = compare(title, ncm_name);
                var p1 = compare(artist, ncm_artist);
                if (p0 == 100 && p1 == 100) {
                    b = k;
                    c = a_k;
                    out[0] = p0;
                    out[1] = p1;
                    break;
                }
                if (p0 > out[0]) {
                    b = k;
                    c = a_k;
                    out[0] = p0;
                } else {
                    if (!artist && (p0 == out[0] && p1 > out[1])) {
                        b = k;
                        c = a_k;
                        out[1] = p1;
                    }
                }
            }
        }
        var res_id = song[b].id;
        var res_name = song[b].name;
        var res_artist = song[b].artists[c].name;
        debug && console(res_id + "-" + res_name + "-" + res_artist);
        //获取封面链接
        infoURL= "http://music.163.com/api/song/detail/?id=" + res_id + "&ids=%5B" + res_id + "%5D";
        try {
            xmlHttp.Open("GET", infoURL, false);
            //noinspection JSAnnotator
            xmlHttp.Option(4) = 13056;
            //noinspection JSAnnotator
            xmlHttp.Option(6) = false;
            xmlHttp.SetRequestHeader("Cookie", "appver=1.5.0.75771");
            xmlHttp.SetRequestHeader("Referer", "http://music.163.com/");
            xmlHttp.SetRequestHeader("Connection", "Close");
            xmlHttp.Send(post_data);
        } catch (e) {
            debug && console("Get Lyric failed");
        }
        if (xmlHttp.Status == 200) {
            var ncm_pic = json(xmlHttp.responseText);
            return ncm_pic.songs[0].album.picUrl;
        }
        }
}
 


function console(s) {
  fb.trace("NCMscript: \n" + s);
}
function del(str, delthis) {
    var s = [str, ""];
    var set = str.indexOf(delthis);

    if (set == -1) {
        return s;
    }
    s[1] = " " + str.substr(set);
    s[0] = str.substring(0, set);

    return s;
}
function compare(x, y) {
    x = x.split("");
    y = y.split("");
    var z = 0;
    var s = x.length + y.length;


    x.sort();
    y.sort();
    var a = x.shift();
    var b = y.shift();

    while (a !== undefined && b !== undefined) {
        if (a === b) {
            z++;
            a = x.shift();
            b = y.shift();
        } else if (a < b) {
            a = x.shift();
        } else if (a > b) {
            b = y.shift();
        }
    }
    return z / s * 200;
}    

function json(text) {
    try {
        var data = JSON.parse(text);
        return data;
    } catch (e) {
        return false;
    }
}

//json2.js
if(typeof JSON!=='object'){JSON={};}
(function(){'use strict';function f(n){return n<10?'0'+n:n;}
    if(typeof Date.prototype.toJSON!=='function'){Date.prototype.toJSON=function(key){return isFinite(this.valueOf())?this.getUTCFullYear()+'-'+
    f(this.getUTCMonth()+1)+'-'+
    f(this.getUTCDate())+'T'+
    f(this.getUTCHours())+':'+
    f(this.getUTCMinutes())+':'+
    f(this.getUTCSeconds())+'Z':null;};String.prototype.toJSON=Number.prototype.toJSON=Boolean.prototype.toJSON=function(key){return this.valueOf();};}
    var cx=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,escapable=/[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,gap,indent,meta={'\b':'\\b','\t':'\\t','\n':'\\n','\f':'\\f','\r':'\\r','"':'\\"','\\':'\\\\'},rep;function quote(string){escapable.lastIndex=0;return escapable.test(string)?'"'+string.replace(escapable,function(a){var c=meta[a];return typeof c==='string'?c:'\\u'+('0000'+a.charCodeAt(0).toString(16)).slice(-4);})+'"':'"'+string+'"';}
    function str(key,holder){var i,k,v,length,mind=gap,partial,value=holder[key];if(value&&typeof value==='object'&&typeof value.toJSON==='function'){value=value.toJSON(key);}
        if(typeof rep==='function'){value=rep.call(holder,key,value);}
        switch(typeof value){case'string':return quote(value);case'number':return isFinite(value)?String(value):'null';case'boolean':case'null':return String(value);case'object':if(!value){return'null';}
            gap+=indent;partial=[];if(Object.prototype.toString.apply(value)==='[object Array]'){length=value.length;for(i=0;i<length;i+=1){partial[i]=str(i,value)||'null';}
                v=partial.length===0?'[]':gap?'[\n'+gap+partial.join(',\n'+gap)+'\n'+mind+']':'['+partial.join(',')+']';gap=mind;return v;}
            if(rep&&typeof rep==='object'){length=rep.length;for(i=0;i<length;i+=1){if(typeof rep[i]==='string'){k=rep[i];v=str(k,value);if(v){partial.push(quote(k)+(gap?': ':':')+v);}}}}else{for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v){partial.push(quote(k)+(gap?': ':':')+v);}}}}
            v=partial.length===0?'{}':gap?'{\n'+gap+partial.join(',\n'+gap)+'\n'+mind+'}':'{'+partial.join(',')+'}';gap=mind;return v;}}
    if(typeof JSON.stringify!=='function'){JSON.stringify=function(value,replacer,space){var i;gap='';indent='';if(typeof space==='number'){for(i=0;i<space;i+=1){indent+=' ';}}else if(typeof space==='string'){indent=space;}
        rep=replacer;if(replacer&&typeof replacer!=='function'&&(typeof replacer!=='object'||typeof replacer.length!=='number')){throw new Error('JSON.stringify');}
        return str('',{'':value});};}
    if(typeof JSON.parse!=='function'){JSON.parse=function(text,reviver){var j;function walk(holder,key){var k,v,value=holder[key];if(value&&typeof value==='object'){for(k in value){if(Object.prototype.hasOwnProperty.call(value,k)){v=walk(value,k);if(v!==undefined){value[k]=v;}else{delete value[k];}}}}
        return reviver.call(holder,key,value);}
        text=String(text);cx.lastIndex=0;if(cx.test(text)){text=text.replace(cx,function(a){return'\\u'+
            ('0000'+a.charCodeAt(0).toString(16)).slice(-4);});}
        if(/^[\],:{}\s]*$/.test(text.replace(/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,'@').replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,']').replace(/(?:^|:|,)(?:\s*\[)+/g,''))){j=eval('('+text+')');return typeof reviver==='function'?walk({'':j},''):j;}
        throw new SyntaxError('JSON.parse');};}}());

