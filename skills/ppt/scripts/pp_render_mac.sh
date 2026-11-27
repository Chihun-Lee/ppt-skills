#!/bin/zsh
# [Mac + Microsoft PowerPoint 전용, 선택] 실제 PowerPoint로 pptx를 열어 PDF로 내보낸다(복구 창이 뜨면 시간 초과). 사용: pp_render.sh <입력.pptx> <출력.pdf>
# Office 샌드박스 권한 창을 피하려고 Group Containers 사본으로 연다. 열려 있는 다른 프레젠테이션은 건드리지 않는다.
G="$HOME/Library/Group Containers/UBF8T346G9.Office"; N="_pp_render_$$.pptx"; cp "$1" "$G/$N"; rm -f "$G/${N:r}.pdf"
osascript <<EOF
with timeout of 420 seconds
tell application "Microsoft PowerPoint"
  open POSIX file "$G/$N"
  set p to presentation "$N"
  save p in POSIX file "$G/${N:r}.pdf" as save as PDF
  close p saving no
end tell
end timeout
EOF
mv "$G/${N:r}.pdf" "$2"; rm -f "$G/$N"
