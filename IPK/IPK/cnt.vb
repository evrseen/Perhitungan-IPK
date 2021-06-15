Namespace htg
    Public Class cnt
        Public Shared Function count(ByVal n1 As Single, n2 As Single, n3 As Single, n4 As Single, n5 As Single, n6 As Single, s1 As Single, s2 As Single, s3 As Single, s4 As Single, s5 As Single, s6 As Single)
            Return ((n1 * s1) + (n2 * s2) + (n3 * s3) + (n4 * s4) + (n5 * s5) + (n6 * s6)) / (s1 + s2 + s3 + s4 + s5 + s6)
        End Function
    End Class
End Namespace