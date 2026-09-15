export function drawTestSample(c,kind){
 if(kind==='current')return;
 const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,1280,720);
 const text=(s,x,y,size=20,color='#344159',family='"Malgun Gothic",sans-serif',weight=400)=>{g.fillStyle=color;g.font=`${weight} ${size}px ${family}`;g.fillText(s,x,y);};
 text(kind==='palette'?'COLOR / CONTRAST':'TYPE / WEIGHT',64,70,16,'#2462e9');
 text(kind==='palette'?'색상과 명암 확인':'글꼴과 굵기 확인',64,125,38,'#18283e',undefined,700);
 if(kind==='palette'){
  const colors=['#FF0000','#FF8000','#FFFF00','#00CC44','#00CCCC','#0066FF','#6633CC','#CC33AA'];
  colors.forEach((hex,i)=>{const x=64+i*146;g.fillStyle=hex;g.fillRect(x,170,132,110);text(hex,x,310,16,'#344159','monospace');for(let j=0;j<4;j++){g.fillStyle=hex;g.fillRect(x,338+j*45,132,40);g.fillStyle=`rgba(255,255,255,${j*.23})`;g.fillRect(x,338+j*45,132,40);}});
  for(let i=0;i<16;i++){g.fillStyle=`rgb(${i*17},${i*17},${i*17})`;g.fillRect(64+i*73,548,73,58);}
  text('BLACK → WHITE',64,641,16,'#344159','monospace');text('고채도 · 밝은 색 · 16단계 회색',800,641,20);
 }else{
  text('글꼴',64,174,17,'#64748b');text('보통 400',355,174,17,'#64748b');text('굵게 700',785,174,17,'#64748b');
  const fonts=[['맑은 고딕','"Malgun Gothic",sans-serif'],['바탕','Batang,serif'],['Arial','Arial,sans-serif'],['Georgia','Georgia,serif']];
  fonts.forEach(([label,family],i)=>{const y=238+i*88;text(label,64,y,22);text('Aa 가나다 0123',355,y,30,'#18283e',family,400);text('Aa 가나다 0123',785,y,30,'#18283e',family,700);g.fillStyle='#e4e8ee';g.fillRect(64,y+28,1150,1);});
  text('크기 비교',64,612,18);[14,18,24,32,42].forEach((size,i)=>text(`${size}px  가Aa`,235+i*190,620,size));
  text('설치되지 않은 글꼴은 대체 글꼴로 표시됩니다.',64,678,16,'#64748b');
 }
}
