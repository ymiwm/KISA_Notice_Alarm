# Todo(Refactory) List
# - KISA(krcert) 사이트 변동 시 코드 조정 필요(ex. ID=="공지")

# ┌────────────────┐
# │     KISA_Notice_Alarm V1.0     │
# │     Author: Alex               │
# │     Last Modified: 2025-01-10  │
# └────────────────┘

echo "
          _____                    _____                    _____                    _____          
         /\    \                  /\    \                  /\    \                  /\    \         
        /::\____\                /::\    \                /::\    \                /::\    \        
       /:::/    /                \:::\    \              /::::\    \              /::::\    \       
      /:::/    /                  \:::\    \            /::::::\    \            /::::::\    \      
     /:::/    /                    \:::\    \          /:::/\:::\    \          /:::/\:::\    \     
    /:::/____/                      \:::\    \        /:::/__\:::\    \        /:::/__\:::\    \    
   /::::\    \                      /::::\    \       \:::\   \:::\    \      /::::\   \:::\    \   
  /::::::\____\________    ____    /::::::\    \    ___\:::\   \:::\    \    /::::::\   \:::\    \  
 /:::/\:::::::::::\    \  /\   \  /:::/\:::\    \  /\   \:::\   \:::\    \  /:::/\:::\   \:::\    \ 
/:::/  |:::::::::::\____\/::\   \/:::/  \:::\____\/::\   \:::\   \:::\____\/:::/  \:::\   \:::\____\
\::/   |::|~~~|~~~~~     \:::\  /:::/    \::/    /\:::\   \:::\   \::/    /\::/    \:::\  /:::/    /
 \/____|::|   |           \:::\/:::/    / \/____/  \:::\   \:::\   \/____/  \/____/ \:::\/:::/    / 
       |::|   |            \::::::/    /            \:::\   \:::\    \               \::::::/    /  
       |::|   |             \::::/____/              \:::\   \:::\____\               \::::/    /   
       |::|   |              \:::\    \               \:::\  /:::/    /               /:::/    /    
       |::|   |               \:::\    \               \:::\/:::/    /               /:::/    /     
       |::|   |                \:::\    \               \::::::/    /               /:::/    /      
       \::|   |                 \:::\____\               \::::/    /               /:::/    /       
        \:|   |                  \::/    /                \::/    /                \::/    /        
         \|___|                   \/____/                  \/____/                  \/____/         

$(Get-Date -Format "yyyy-MM-dd(dddd) HH:mm:ss")
Program:`tKISA 공지 알람
Owner:`t`t하나금융TI/SK Shieldus
"

$second=-1
$kisa_latest_id=0
$kisa_cur_id=0
$kisa_cur_title=""
$flag=$false
$count=0

while($true){
    if($(Get-Date).Second -ne 0){
        $second=$(Get-Date).Second
        Write-Progress -Activity "공지 확인 대기" -Status "$(59-$second)초 후 공지 확인" -PercentComplete $($second*100/59)
        continue
    }
    Write-Progress "공지 확인 중" -completed
    $cur_time=Get-Date
    $count+=1

    try{
        $kisa_html=Invoke-WebRequest "https://krcert.or.kr/kr/bbs/list.do?menuNo=205020&bbsId=B0000133"
        $kisa_cur_id=$kisa_html.AllElements | ?{$_.Class -eq "num"} | select -index 1 | select -ExpandProperty InnerText
        $kisa_cur_title=$kisa_html.AllElements | ?{$_.Class -eq "sbj tal"} | select -index 0 | select -ExpandProperty InnerText
        if($kisa_cur_id -eq "공지"){
            $kisa_cur_id=$kisa_html.AllElements | ?{$_.Class -eq "num"} | select -index 2 | select -ExpandProperty InnerText
            $kisa_cur_title=$kisa_html.AllElements | ?{$_.Class -eq "sbj tal"} | select -index 1 | select -ExpandProperty InnerText
        }
    }
    catch{
        echo "[-] ($($cur_time.ToString("HH:mm"))) 오류 발생.($_)"
    }

    if($kisa_latest_id -ne $kisa_cur_id){
        $diff=$kisa_cur_id-$kisa_latest_id
        if($diff -eq $kisa_cur_id){
            echo "[+] ($($cur_time.ToString("HH:mm"))) 최초 확인용 공지"
        }
        else{
            echo "[+] ($($cur_time.ToString("HH:mm"))) 신규 공지 발생.(신규 공지 건수: $diff 건)"
        }
        echo "`t($kisa_cur_id) $kisa_cur_title"
        $kisa_latest_id=$kisa_cur_id
        $flag=$true
    }
    
    if($flag){
        $wscript=New-Object -ComObject Wscript.Shell
        $wscript.Popup("($kisa_cur_id) $kisa_cur_title",0,"KISA 최신 공지",0x0) > $null
        $flag=$false
    }
}
