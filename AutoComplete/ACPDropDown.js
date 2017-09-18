var cD = gid('completeDiv');
var tD = gid('titleDiv');
var oD = gid('holderDiv');
var cS = false;
var aS = -1;
var tL = 0;
var lS = ''
var lA = null;

function gid(e){
	return document.getElementById(e);
}
function keyComplete(e, V, ACP){
	switch(e.keyCode){
		case 38:
			aS-=2;
		case 40:
			if(!cS) showComplete()
			aS+=1;
			aS=aS<0?0:aS>tL?tL:aS;
			var nC=gid('completeItem'+aS);
			hl(nC, ACP);
			if(cD.scrollTop+parseInt(cD.style.height)<nC.offsetTop+nC.offsetHeight)
				cD.scrollTop=nC.offsetTop-parseInt(cD.style.height)+nC.offsetHeight;
			if(cD.scrollTop>nC.offsetTop) cD.scrollTop=nC.offsetTop;
			return false;
		case 13:
			Choose(gid('completeItem'+aS), ACP, V);
			return false;
	}
	return true;
}
function lV(V, ACP){
	if(event.keyCode<=42 && event.keyCode!=0 && event.keyCode!=8 && event.keyCode!=32 ) return;
	prepACP(ACP);
	if(V.value=='') return;
	if(V.value==lS){showComplete(); return; }
	if(V.value!=lS){
		var ary1 = new Array();
		var ary2 = new Array();
		var ary3 = new Array(ACP.fields[0].length);
		for(var i=0; i<ary3.length; i++) ary3[i]=0;

		for(var i=0, j=0; i<ACP.AutoCompleteOn.length; i++){
			var cFl  = ACP.AutoCompleteOn[i];
			var cAry = ACP.fields[cFl];
			var lD   = 0;
			var hD   = cAry.length;
			while (lD <= hD){
				mD = Math.floor((lD + hD) / 2);
				if (V.value.toLowerCase() <= (cAry[ACP.hlpSortAry[cFl][mD]]+'').toLowerCase()){
					hD = mD - 1;
				} else {;
					lD  = mD + 1;
				}
			}
			while(lD<cAry.length && (cAry[ACP.hlpSortAry[cFl][lD]]+'').toLowerCase().indexOf(V.value.toLowerCase())==0){
				if(ary3[ACP.hlpSortAry[cFl][lD]]==0){
					nVl = ACP.fields[ACP.sF][ACP.hlpSortAry[cFl][lD]];
					if(typeof(nVl)=='string'){
						nVl = nVl.toLowerCase();
						if(nVl.indexOf('<')>0) nVl = nVl.replace(/<.*?>/gi, '');
					}
					ary1[j++]={ix:ACP.hlpSortAry[cFl][lD], vl:nVl};
					ary3[ACP.hlpSortAry[cFl][lD]]=1;
				}
				lD++;
			}
		}


		for(var i=0, j=0; i<ACP.fieldW.length; i++) if(ACP.fieldW[i]>0) j++;
		if(ACP.sD){quick(ary1, 1)} else {quick(ary1, 0)}
		for(var i=0; i<ary1.length; i++) ary2[i] = ary1[i].ix;
		ACP.numFields=j;
		tL = ary2.length-1;
		lA = ary2;
		lS = V.value;
	} else {
		ary2 = lA;
	}
	dT(V, ary2, ACP);
}
var ix=0;
function dH(ACP, V){
	nO = new Array();
	x  = 0;
	for(var j=0; j<ACP.fields.length; j++){
		if(ACP.fieldW[j]>0){
			w=(((parseInt(tD.style.width))*(ACP.fieldW[j]/100))-(18/ACP.numFields));
			w=(w<0)?0:w+(j+1==ACP.fields.length?10:0);
			nO[x++]='<span style="width:'+w+'px;cursor:hand;" onmousedown="orderComplete('+j+', gid(\''+V.id+'\'), '+ACP.Nm+');">'+ACP.Nms[j]+(ACP.sF==j?'<span style="width:16px;background:url(grid.png) -'+(ACP.sD?20:40)+'px 50% no-repeat;"></span>':'')+'</span>';
		}
	}
	return nO.join('');
}
function dT(V, ary, ACP){
	cD.innerHTML='';
	cD.hC='0';
	if(ary.length>0){
		cD.style.width      = tD.style.width = typeof(ACP.Width)!='undefined'?ACP.Width:V.offsetWidth;
		tD.style.top        = V.offsetTop+V.offsetHeight-1;
		cD.style.left       = tD.style.left  = V.offsetLeft;
		tD.innerHTML=dH(ACP, V);

		nO = new Array();
		x  = 0;
		cD.style.top             = tD.offsetTop+tD.offsetHeight;
		cD.style.height          = null;
		cD.style.backgroundColor = ACP.bgColor;
		nO[x++]='<table border="0" cellspacing="0" cellpadding="0" width="'+(parseInt(cD.style.width)-20)+'">';
		for(var i=0; i<ary.length; i++){
			nO[x++]='<tr id="completeItem'+i+'" style="cursor:hand;" onmousedown="Choose(this, '+ACP.Nm+', gid(\''+V.id+'\'));" key="'+ary[i]+'" onmouseover="hl(this, '+ACP.Nm+');">';
			for(var j=0; j<ACP.fields.length; j++){
				if(ACP.fieldW[j]>0){
					w=(((parseInt(cD.style.width))*(ACP.fieldW[j]/100))-(20/ACP.numFields));
					w=(w<0)?0:w;
					nO[x++]='<td nowrap width="'+ACP.fieldW[j]+'%"><div style="'+ACP.Styles[j]+';width:'+w+'px;">'+ACP.fields[j][ary[i]]+'</div></td>';
				}
			}
			nO[x++]='</tr>';
		}
		nO[x++]='</table>';

		cD.innerHTML=nO.join('');
		cD.hC='1';
		if (cD.offsetHeight>ACP.maxHeight) cD.style.height=ACP.maxHeight+'px';
	}
	showComplete();
}
var lastE=null;
var xi=0;
function hl(e, ACP){
	aS=parseInt(e.id.replace(/completeItem/, ''));
	if(lastE!=null) uhl(lastE, ACP);
	for(var j=0; j<e.children.length; j++){
		fTD=e.children[j];
		for(var i=0; i<fTD.children.length; i++){
			fDiv=fTD.children[i];
			fDiv.outerHTML=fDiv.outerHTML.replace(/style=\".*?\"/gi, 'style="'+ACP.Styles[j].replace(/(background-)?color:.*?;/gi, '')+';'+ACP.HLStyle+';width:'+fDiv.style.width+';"');
		}
	}
	lastE = e;
}
function uhl(e, ACP){
	for(var j=0; j<e.children.length; j++){
		fTD=e.children[j];
		for(var i=0; i<fTD.children.length; i++){
			fDiv=fTD.children[i];
			fDiv.outerHTML=fDiv.outerHTML.replace(/style=\".*?\"/gi, 'style="'+ACP.Styles[j]+';width:'+fDiv.style.width+';"');
		}
	}
	lastE = null;
}
function Choose(e, ACP, V){
	lS=V.value=ACP.fields[ACP.TextVal][parseInt(e.key)];
	//gid(V.id+'Val').value=ACP.fields[ACP.ValueOf][parseInt(e.key)];
	cS = false;
	hideComplete();
}
function hideComplete(){
	if (cS) return;
	cD.style.visibility = tD.style.visibility = 'hidden';
}
function showComplete(){
	if(cD.hC=='0'){
		cS=false;
		hideComplete();
	} else {
		cS=true;
		cD.style.zIndex=10;
		cD.style.visibility = tD.style.visibility = '';
	}
}
function offComplete(e){
	cS = false;
	setTimeout('hideComplete()', typeof(e)!='undefined'?e:500);
}
function onComplete(){
	cS = true;
}
function orderComplete(e, V, ACP){
	cS=true;
	lS='';
	if(ACP.sF==e){
		ACP.sD=!ACP.sD;
	} else {
		ACP.sF=e;
		ACP.sD=true;
		n='';
		for(var i=0; i<ACP.hlpSortAry[e].length; i++) ACP.sortAry[i] = ACP.hlpSortAry[e][i];
	}
	lV(V, ACP);
}
function prepACP(ACP){
	aS = -1;
	if(typeof(ACP.sortAry)!='undefined') return;
	if(typeof(ACP.OrderBy)=='undefined') ACP.OrderBy=-1;

	ACP.sortAry    = new Array();
	ACP.hlpSortAry = new Array(ACP.fields.length);
	for(var i=0; i<ACP.fields.length; i++) ACP.hlpSortAry[i] = new Array();

	ACP.sD = ACP.hsD = ACP.hsF;
	ACP.sD = true;
	ACP.sF = 0;
	for(var i=0; i<ACP.fields[0].length; i++){
		ACP.sortAry[i]=i;
		for(var j=0; j<ACP.hlpSortAry.length; j++) ACP.hlpSortAry[j][i] = i;
	}
	for(var i=0; i<ACP.fields.length; i++){
		if(i!=ACP.OrderBy){
			if(ACP.fieldW[i]>0){
				tmpAry    = new Array();
				for(var j=0; j<ACP.fields[i].length; j++){
					nVl = ACP.fields[i][j].toLowerCase();
					if(typeof(nVl)=='string'){
						nVl = nVl.toLowerCase()
						if(nVl.indexOf('<')>0) nVl = nVl.replace(/<.*?>/gi, '');
					}
					tmpAry[j] = {ix:j, vl:nVl};
				}
				quick(tmpAry, 1);
				for(var j=0; j<tmpAry.length; j++) ACP.hlpSortAry[i][j] = tmpAry[j].ix;
			}
		}
	}
}
function completePrep(e, ACP){
	e.outerHTML=e.outerHTML.substr(0, e.outerHTML.length-1)+' onkeydown="keyComplete(event, this, '+ACP+')" onkeyup="lV(this, '+ACP+')" onmouseover="this.onkeyup();" onmouseout="offComplete();" onmousemove="onComplete();" onfocus="this.onkeyup();" onkeyup="keyComplete()" onblur="cS=false;offComplete(50);">';
}

function qsCmp(a, b, d){
	return a.vl>b.vl?(d?1:0):(d?0:1);
}
function Quicksort(vec, loBound, hiBound, d)
{
	var pivot, pivot2, loSwap, hiSwap, temp, temp2;

	if (hiBound - loBound == 1){
		if (qsCmp(vec[loBound], vec[hiBound], d)){
			temp  = vec[loBound];
			vec[loBound] = vec[hiBound];
			vec[hiBound] = temp;
		}
		return;
	}

	pivot  = vec[parseInt((loBound + hiBound) / 2)];
	vec[parseInt((loBound + hiBound) / 2)] = vec[loBound];
	vec[loBound] = pivot;
	loSwap = loBound + 1;
	hiSwap = hiBound;

	while (loSwap < hiSwap){
		while (loSwap <= hiSwap && !qsCmp(vec[loSwap], pivot, d)) loSwap++;
		while (qsCmp(vec[hiSwap], pivot, d) && hiSwap>=loSwap)hiSwap--;
		if (loSwap < hiSwap){
			temp  = vec[loSwap];
			vec[loSwap] = vec[hiSwap];
			vec[hiSwap] = temp;
		}
	}

	vec[loBound] = vec[hiSwap];
	vec[hiSwap] = pivot;
	if (loBound < hiSwap - 1) Quicksort(vec, loBound, hiSwap - 1, d);
	if (hiSwap + 1 < hiBound) Quicksort(vec, hiSwap + 1, hiBound, d);
}
function quick(n, d){
	if (typeof(d)=='undefined') d=true;
	Quicksort(n, 0, n.length-1, d);
}