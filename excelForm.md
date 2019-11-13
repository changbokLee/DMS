## 라이브러리 마다 지원 형식폼이 다르다
> openpyxl xlsx,.xlsm,.xltx,.xltm을 지원한다. => 스타일 지정할때 편하다
  XlsWriter(쓰기) xlsx지원 
  xlrd , xlwt 
  Pandas로도 가능하다. =>(엑셀설치 필요없음 , 윈도우 ,MacOs, Linux) => 엑셀의 값을 병합하는 데이터분석할때 좋다
  <hr>
  xlwings(윈도우, MacOs) = 엑셀매크로 자동화,
  pywin32(윈도우) => 설치된 엑셀이 필요
### 라이브러리마다  지원형식이 안되는 폼이 몇가지가 있다.
> openpyxl => 테이블 테마 플랫폼이 지원이 안됨 ,

### 데이터 병합방법
> 각자 다른 xlsx파일이 있다면
padnas => df_삼성전자 = pd.read_excel~~~
빈데이터 프레임가지고 나머지 데이터 프레임을 토탈해서 더한다
df_merge =~~~ 삼성
df_merge ~~~~ LG



