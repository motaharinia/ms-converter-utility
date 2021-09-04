## ms-converter-utility
this project is a shared common utility project for microservice projects as a pom dependency to reduce boilerplate code in multiple projects.

### Project Descriptions:
build with java-doc:
- if your windows account is for example "myuser" and you have "C:\Users\MyUser\.m2\repository\com\motaharinia\ms-converter-utility" folder, remove it.
- open ms-converter-utility project by IntellliJ IDEA(enable auto import , Windows defender automatic fix) and check in "project structure" that jdk 11 is selected.
- in IntellliJ IDEA open right side maven panel and select lifecycle/install and run maven to build it in "C:\Users\MyUser\.m2\repository\com\motaharinia\ms-converter-utility"

### Change Log:
- 0.0.1 : initial project
- 0.0.2 : add type conversion for Instant,LocalDate,LocalDateTime in CustomMapper class. add languages for customvalidation messages

### IntellliJ IDEA Configurations:
- IntelijIDEA: Help -> Edit Custom Vm Options -> add these two line:
    - -Dfile.encoding=UTF-8
    - -Dconsole.encoding=UTF-8
- IntelijIDEA: File -> Settings -> Editor -> File Encodings-> Project Encoding: form "System default" to UTF-8. Maybe it affected somehow.
- IntelijIDEA: File -> Settings -> Editor -> General -> Code Completion -> check "show the documentation popup in 500 ms"
- IntelijIDEA: File -> Settings -> Editor -> General -> Auto Import -> check "Optimize imports on the fly (for current project)"
- IntelijIDEA: File -> Settings -> Editor -> Color Scheme -> Color Scheme Font -> Scheme: Default -> uncheck "Show only monospaced fonts" and set font to "Tahoma"
- IntelijIDEA: File -> Settings -> Editor -> Proofreading -> add "farsi" language for graze spell check
- IntelijIDEA: Run -> Edit Configuration -> Spring Boot -> XXXApplication -> Configuration -> Environment -> VM Options: -Dspring.profiles.active=dev
- IntelijIDEA: Run -> Edit Configuration -> Spring Boot -> XXXApplication -> Code Coverage -> Fix the package in include box

<hr/>
<a href="mailto:eng.motahari@gmail.com?"><img src="https://img.shields.io/badge/gmail-%23DD0031.svg?&style=for-the-badge&logo=gmail&logoColor=white"/></a>


