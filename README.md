# reset-password-active-directory

**Active Directory 사용자의 비밀번호를 변경**하는 **응용 프로그램**입니다.

<small>[이전글: Active Directory 비밀번호 변경](/active-directory-change-password/)의 코드를 참조해서 작업하시면서 잘 안되는 부분이 있다고 하여 다시 작성합니다.</small>


## 사용

![active-directory-user-change-password](/img/active-directory-user-change-password.png)


1. 도메인을 입력합니다. e.g.) bbon.kr
2. Active Directory 관리자 계정 이름을 입력합니다. e.g.) Administrator
3. Active Directory 관리자 계정 비밀번호를 입력합니다.
4. 비밀번호를 변경할 계정 이름을 입력합니다.
5. Connect 버튼을 클릭합니다.
6. 연결에 문제가 없는 경우 변경할 비밀번호 입력 텍스트 박스가 활성화됩니다.
7. 변경할 비밀번호, 비밀번호 확인 텍스트 박스에 변경할 비밀번호를 동일하게 입력합니다.
8. Change 버튼을 클릭합니다. 

> * Active Directory 관리자 계정 정보 <small>계정이름, 비밀번호</small>가 필요합니다.
> * .NET Framework 4.0이 필요합니다.