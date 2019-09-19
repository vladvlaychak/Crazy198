DIm a()
DIm b()
Dim c()
Redim a(3)
a(0)=2
a(1)=67
a(2)=21
a(3)=14
Redim b(3)
b(0)=7
b(1)=62
b(2)=2
b(3)=14
Redim c(3)
c(1...3)= a(1...3)+b(1...3)
MsgBox c(3)