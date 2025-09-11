// ------------------ Fonts ------------------
const fonts = ["Arial","Times New Roman","Courier New","Verdana","Georgia","Impact","Comic Sans MS","Trebuchet MS","Tahoma",
"Lucida Console","Palatino Linotype","Segoe UI","Roboto","Open Sans","Lato","Montserrat","Oswald","Raleway",
"Merriweather","PT Sans","Ubuntu","Droid Sans","Source Sans Pro","Noto Sans","Playfair Display","Fira Sans",
"Arimo","Roboto Slab","Poppins","Nunito","Josefin Sans","Roboto Mono","Lora","Cairo","Mukta","Karla","Work Sans",
"Exo 2","IBM Plex Sans","Cabin","Inconsolata","Quicksand","Anton","Bebas Neue","Dancing Script","Pacifico",
"Permanent Marker","Shadows Into Light","Amatic SC","Chewy","Fredericka the Great","Caveat","Indie Flower",
"Satisfy","Gloria Hallelujah","Baloo 2","Fredoka One","Bangers","Hind","Catamaran","Crimson Text","Heebo",
"Overpass","Manrope","Signika","Titillium Web","Varela Round","Kanit","Barlow","Nanum Gothic","Nanum Myeongjo",
"Gothic A1","Rubik","Prompt","Sarabun","Trirong","Teko","Public Sans","Space Grotesk","Be Vietnam","Asap",
"Chivo","Exo","Fira Sans Condensed","Hind Madurai","Inter","Lexend","Noto Serif","Red Hat Display","Red Hat Text",
"Spectral","Work Sans","PT Serif","Vollkorn","Merriweather Sans","Hepta Slab"];
const fontSelect = document.getElementById("fontSelect");
fonts.forEach(f => { const o=document.createElement("option"); o.value=f;o.innerText=f;fontSelect.appendChild(o); });

// ------------------ Ribbon Tabs ------------------
function switchTab(tab){
    document.querySelectorAll('.tabPanel').forEach(p=>p.style.display='none');
    document.querySelectorAll('.tabButton').forEach(b=>b.classList.remove('active'));
    document.getElementById(tab+'Tab').style.display='flex';
    document.querySelector(`.tabButton[onclick="switchTab('${tab}')"]`).classList.add('active');
}

// ------------------ Editor Commands ------------------
function execCmd(command,value=null){
    const editor=document.querySelector('.page:last-child .editor');
    editor.focus();
    document.execCommand(command,false,value);
}

// ------------------ Custom Fonts ------------------
function addCustomFont(){
    const f=document.getElementById('customFontInput').value.trim(); 
    if(!f) return;
    const link=document.createElement('link'); 
    link.href=`https://fonts.googleapis.com/css2?family=${f.replaceAll(' ','+')}&display=swap`; 
    link.rel='stylesheet'; document.head.appendChild(link);
    const o=document.createElement('option'); o.value=f;o.innerText=f;fontSelect.appendChild(o);
    execCmd('fontName',f);
}

// ------------------ Tables & Images ------------------
function insertTable(){const rows=prompt("Rows",2),cols=prompt("Columns",2); if(!rows||!cols) return;
let table="<table border='1' style='border-collapse: collapse;'>"; for(let r=0;r<rows;r++){table+="<tr>"; for(let c=0;c<cols;c++) table+="<td>&nbsp;</td>"; table+="</tr>";} table+="</table>";
execCmd('insertHTML',table);}
function insertImage(){const url=prompt("Image URL:"); if(url) execCmd('insertImage',url);}

// ------------------ Page Layout ------------------
function setPageOrientation(v){document.querySelectorAll('.page').forEach(p=>{if(v==='portrait'){p.style.width="21cm";p.style.height="29.7cm";}else{p.style.width="29.7cm";p.style.height="21cm";}});}
function zoomIn(){document.querySelectorAll('.page').forEach(p=>p.style.transform="scale(1.2)");}
function zoomOut(){document.querySelectorAll('.page').forEach(p=>p.style.transform="scale(1)");}

// ------------------ Editor Pages ------------------
const editorWrapper=document.getElementById('editorWrapper');
function createPage(content=""){
    const page=document.createElement('div'); page.className='page';
    const editor=document.createElement('div'); editor.className='editor'; editor.contentEditable=true; editor.innerHTML=content; editor.addEventListener('input',checkPageOverflow);
    page.appendChild(editor);
    editorWrapper.appendChild(page);
}
function checkPageOverflow(){ const page=document.querySelector('.page:last-child'); const editor=page.querySelector('.editor'); if(editor.scrollHeight>editor.clientHeight){createPage();}}

// ------------------ File Import ------------------
function importFile(event){
    const file = event.target.files[0];
    if(!file) return;

    let ext = file.name.split('.').pop().toLowerCase();

    const hMap = {
        'hoc':'docx','hocx':'docx',
        'hot':'dotx','hotx':'dotx',
        'hocm':'docm','hotm':'dotm',
        'htf':'xml','hxt':'xml',
        'hml':'html','hht':'html','hhtml':'html'
    };
    if(hMap[ext]) ext = hMap[ext];

    const reader = new FileReader();
    reader.onload = async function(e){
        const content = e.target.result;
        editorWrapper.innerHTML = '';
        createPage();
        const editor = document.querySelector('.page:last-child .editor');

        try{
            if(['txt','xml','hoc','hocx','hot','hotx','hocm','hotm','htf','hxt'].includes(ext)){
                editor.innerText = content;
            } else if(['html','hml','hht','hhtml','mht','mhtml'].includes(ext)){
                editor.innerHTML = content;
            } else if(['doc','dot','docx','dotx','docm','dotm'].includes(ext)){
                const arrayBuffer = content;
                const result = await mammoth.extractRawText({arrayBuffer});
                editor.innerText = result.value;
            } else if(['rtf'].includes(ext)){
                if(typeof RTFJS!=='undefined'){
                    const doc = new RTFJS.Document(content);
                    const rendered = await doc.render();
                    editor.innerHTML = '';
                    editor.appendChild(rendered);
                } else {
                    editor.innerText = content;
                }
            } else {
                editor.innerText = content;
            }
        } catch(err){
            alert("Error loading file: "+err.message);
        }
    }

    if(['doc','dot','docx','dotx','docm','dotm'].includes(ext)){
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

// ------------------ File Save ------------------
function saveFile(){
    const editor=document.querySelector('.page:last-child .editor');
    const content=editor.innerHTML;
    const fn=prompt("File name with extension","document.txt"); if(!fn) return;
    const b=new Blob([content],{type:"text/html"}); const a=document.createElement('a'); a.href=URL.createObjectURL(b); a.download=fn; a.click();
}

// ------------------ Save As with dropdown ------------------
const saveFormatSelect = document.getElementById('saveFormat');
const formats = [
  "txt","doc","dot","docx","dotx","docm","dotm","rtf","xml",
  "mht","mhtml",
  "hoc","hocx","hot","hotx","hocm","hotm","htf","hxt","hml","hht","hhtml"
];
formats.forEach(f => { const o = document.createElement('option'); o.value = f; o.innerText = f; saveFormatSelect.appendChild(o); });

function saveAs(){
    const editor = document.querySelector('.page:last-child .editor');
    const content = editor.innerHTML;
    const ext = saveFormatSelect.value;
    if(!ext) return;
    let fn = prompt("File name (without extension):","document");
    if(!fn) return;
    fn = fn + "." + ext.toLowerCase();

    let mime = "text/html"; 
    if(["txt","rtf","xml","mht","mhtml","hoc","hocx","hot","hotx","hocm","hotm","htf","hxt"].includes(ext)) mime = "text/plain";

    const blob = new Blob([content],{type:mime});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fn;
    a.click();
}

// ------------------ Undo/Redo ------------------
function undo(){execCmd('undo');}
function redo(){execCmd('redo');}

// ------------------ Find/Replace ------------------
function findReplace(){const f=prompt("Find:"),r=prompt("Replace with:"); if(!f) return; document.querySelectorAll('.editor').forEach(e=>{e.innerHTML=e.innerHTML.split(f).join(r||'');});}

// ------------------ Main Screen ------------------
const templates={blank:{content:''}};
function newDocument(type='blank'){document.getElementById('mainScreen').style.display='none'; document.getElementById('editorScreen').style.display='block'; editorWrapper.innerHTML=''; createPage(templates[type].content);}
