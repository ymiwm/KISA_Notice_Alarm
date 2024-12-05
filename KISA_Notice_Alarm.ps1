# ┌────────────────────────────────┐
# │     KISA_Notice_Alarm V1.0     │
# │     Author: Alex               │
# │     Last Modified: 2024-12-03  │
# └────────────────────────────────┘

# Todo List
# 1. MessageBox와 Non-Block하게 코드 실행
# 2. KISA(krcert) 사이트 변동 시 코드 조정 필요 -> Class, Tag로 고정하거나 "공지" 행 제외

$kisa_latest_id=0
$kisa_cur_id=0
$kisa_cur_title=""
$flag=$false
$count=0

echo "KISA_Notice_Alarm V1.0"
echo "시작 시간: $($(Get-Date).ToString('yyyy-MM-dd hh:mm:ss'))"
echo "================================================================"

while($true){
    if($(Get-Date).Second -ne 0){
        continue
    }
    $cur_time=Get-Date
    $count+=1

    $kisa_html=Invoke-WebRequest "https://krcert.or.kr/kr/bbs/list.do?menuNo=205020&bbsId=B0000133"

    try{
        $kisa_cur_id=$kisa_html.AllElements | ?{$_.Class -eq "num"} | select -index 1 | select -ExpandProperty InnerText
        $kisa_cur_title=$kisa_html.AllElements | ?{$_.Class -eq "sbj tal"} | select -index 0 | select -ExpandProperty InnerText
    }
    catch{
        echo "[-] ($($cur_time.ToString("HH:mm")) HTML 파싱 오류 발생.(Count: $count)($_)"
    }

    if($kisa_latest_id -ne $kisa_cur_id){
        echo "[+] ($($cur_time.ToString("HH:mm"))) KISA 신규 공지 발생.(Count: $count)"
        echo "`t($kisa_cur_id) $kisa_cur_title"
        $kisa_latest_id=$kisa_cur_id
        $flag=$true
    }
    else{
        echo "[+] ($($cur_time.ToString("HH:mm"))) KISA 신규 공지 없음.(Count: $count)"
        $flag=$false
    }
    
    if($flag){
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("($kisa_cur_id) $kisa_cur_title",0,"KISA 신규 공지 발생",0x0) > $null
    }
}
