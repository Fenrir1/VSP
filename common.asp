<%

function Max(s, i)
  if s <= i then 
		Max=i
	else 
		Max=s
  end if
  
end function

function CutStr(s, i, c)
  if i < 1 then i=1 end if
  CutStr=left(s, i-1)+mid(s, i+c)
end function

function ZERO(v)
  ZERO="0"
  if Len(v) > 0 then ZERO=v end if
end function

function NBSP(v)
  NBSP="&nbsp;"
  if Len(v) > 0 then NBSP=v end if
end function

function IIF(expr, t, f)
  if expr = true then IIF=t else IIF=f end if
end function

function DateTimeFormat(t, Stencil)
  if Len(t) > 0 then
    dd=CStr(Day(t))
    if Len(dd) = 1 then dd="0"&dd end if
    mm=CStr(Month(t))
    if Len(mm) = 1 then mm="0"&mm end if
    yy=CStr(Year(t))
    if (Len(yy) = 4) and (InStr(Stencil, "yyyy") = 0) then yy=Right(yy, 2) end if
    hh=CStr(Hour(t))
    if Len(hh) = 1 then hh="0"&hh end if
    nn=CStr(Minute(t))
    if Len(nn) = 1 then nn="0"&nn end if
    ss=CStr(Second(t))
    if Len(ss) = 1 then ss="0"&ss end if

    if Stencil <> "" then
      if InStr(Stencil, "yyyy") > 0 then Stencil=replace(Stencil, "yyyy", yy) end if
      if InStr(Stencil, "yy") > 0 then Stencil=replace(Stencil, "yy", yy) end if
      if InStr(Stencil, "mm") > 0 then Stencil=replace(Stencil, "mm", mm) end if
      if InStr(Stencil, "dd") > 0 then Stencil=replace(Stencil, "dd", dd) end if
      if InStr(Stencil, "hh") > 0 then Stencil=replace(Stencil, "hh", hh) end if
      if InStr(Stencil, "nn") > 0 then Stencil=replace(Stencil, "nn", nn) end if
      if InStr(Stencil, "ss") > 0 then Stencil=replace(Stencil, "ss", ss) end if
      DateTimeFormat=Stencil
    else
      DateTimeFormat = dd & "." & mm & "." & yy & " " & hh & ":" & nn
    end if
  else
    DateTimeFormat = "&nbsp;"
  end if
end function
%>
