# akram-alazani.github.io

اكرم العزاني — نظام إدارة متكامل (Microsoft Access)

هذا المستودع مخصص لمشاركة ملفات مشروع نظام إدارة متكامل مبني بـ Microsoft Access ويتضمن موديولات VBA، نماذج (Forms)، تقارير، استعلامات، وجداول لإدارة:

- المبيعات
- المشتريات
- المخزون
- العملاء
- الموردين
- الموظفين
- الرواتب
- الحضور والانصراف
- المستخدمين والصلاحيات

ملاحظة: حاليًا يوجد في المستودع ملف README فقط. لتمكين المراجعة الهندسية وتحسين النظام أحتاج ملفات المشروع الكاملة.

ما أحتاجه منك لتبدأ المراجعة الشاملة:

1. ملف قاعدة البيانات (.accdb أو .mdb) — الخيار الموصى به. ضع الملف داخل مجلد `database/` بالمستودع، أو ارفعه هنا في المحادثة.

    مثال مسار مقترح: `database/MySystem.accdb`

2. إن لم ترغب بمشاركة ملف القاعدة الكامل، صِدّر الكائنات (Forms, Reports, Modules, Queries, Macros) كملفات نصية وارفَعها داخل مجلد `VBA/` أو `export/` بالمستودع.

    دليل سريع لتصدير كل الكائنات باستخدام VBA (شغّل داخل قاعدة Access):

```vb name=Export_All_Access_Objects.bas
Sub ExportAllObjects()
    Dim exportPath As String
    exportPath = "C:\ExportAccessObjects\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    Dim ao As AccessObject
    ' Export Forms
    For Each ao In CurrentProject.AllForms
        Application.SaveAsText acForm, ao.Name, exportPath & "Form_" & ao.Name & ".txt"
    Next ao
    ' Export Reports
    For Each ao In CurrentProject.AllReports
        Application.SaveAsText acReport, ao.Name, exportPath & "Report_" & ao.Name & ".txt"
    Next ao
    ' Export Modules (standard & class)
    For Each ao In CurrentProject.AllModules
        On Error Resume Next
        Application.SaveAsText acModule, ao.Name, exportPath & "Module_" & ao.Name & ".txt"
        On Error GoTo 0
    Next ao
    ' Export Queries
    For Each ao In CurrentProject.AllQueries
        Application.SaveAsText acQuery, ao.Name, exportPath & "Query_" & ao.Name & ".txt"
    Next ao
    ' Export Macros
    For Each ao In CurrentProject.AllMacros
        Application.SaveAsText acMacro, ao.Name, exportPath & "Macro_" & ao.Name & ".txt"
    Next ao

    MsgBox "Export finished to: " & exportPath
End Sub
```

ملاحظات أمنية قبل الرفع:
- تَأكَد من إزالة أو تعمية أي بيانات حساسة (مثل أرقام بطاقات ائتمان أو بيانات شخصية غير لازمة).
- إن أردت، أرسل نسخة اختبارية (عينة بيانات) بدل القاعدة الحقيقية.

ما سأقوم به بعد استلام الملفات:
- استخراج كل الأكواد والكائنات والجداول والعلاقات.
- إجراء مراجعة هندسية كاملة: اكتشاف الأخطاء، اقتراح تحسينات الأداء، إزالة التكرار، وإعادة تنظيم الأكواد.
- تحسينات أمنية: تشفير كلمات المرور، نظام صلاحيات متعدد المستويات، منع الوصول ال��باشر للجداول.
- تنفيذ ميزات مطلوبة: سجل دخول المستخدمين، نظام نسخ احتياطي واستعادة، لوحة تحكم (Dashboard)، تنبيهات مخزون، دعم الباركود، تقارير وفواتير احترافية.
- تزويدك: الأكواد المحسنة (VBA)، SQL للجداول الجديدة، شرح التعديلات، وخطة نشر جاهزة للإنتاج.

هل تريد أن أُنشئ فورًا ملفات هيكل (SQL) مقترحة للجداول الجديدة ونموذج لواجهة Dashboard قبل استلام القاعدة؟ أم تفضّل أن أبدأ مباشرةً بعد رفع الملف؟

---

License: MIT

Contact: r1gb (GitHub)
