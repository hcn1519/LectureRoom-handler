# LectureRoom-handler
서강대학교 강의실 정보를 정리하여 보여주는 파이썬 코드입니다.

#### 원본 파일 - 서강대학교 개설 교과목 정보에서 다운 받을 수 있습니다.
<img src="https://rawgit.com/hcn1519/LectureRoom-handler/master/images/image1.png">

<br/>

#### 원본 파일에서 필요한 정보만 남겨두고 나머지 칼럼들은 제거
<img src="https://rawgit.com/hcn1519/LectureRoom-handler/master/images/image2.png">

<ol>
<li>wb = openpyxl.load_workbook('/Users/changnam/Desktop/lectureSeed.xlsx')의 경로를 알맞게 설정합니다.</li>
<li>프로그램을 돌립니다.</li>
</ol>

```
  python makeExcelForLecture.py
```


#### 결과물1: lectureRoom.xlsx, 파이썬 정규표현식을 사용하여 강의실 번호로 데이터 재구성
<img src="https://rawgit.com/hcn1519/LectureRoom-handler/master/images/image3.png">

<br/>

#### 결과물2: lectureTime.xlsx, 파이썬 정규표현식을 사용하여 강의시간으로 데이터 재구성
<img src="https://rawgit.com/hcn1519/LectureRoom-handler/master/images/image4.png">
