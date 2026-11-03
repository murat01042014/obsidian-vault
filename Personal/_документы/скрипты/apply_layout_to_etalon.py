# Применяет к эталону ИТД раскладку из build_itd_final.py: отступы Приложения 1 (по образцу Мурата), подписи, уменьшенные
# отступы в теле и Приложении 2, отступ номер-текст 10 / 12,7 мм, запрет разрыва блоков подписей (Мурат, 03.11.2026).
import importlib.util, zipfile, os, sys
spec = importlib.util.spec_from_file_location("b", r"D:\Кадры\_исходники\build_itd_final.py")
b = importlib.util.module_from_spec(spec); spec.loader.exec_module(b)
p = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(p); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
with zipfile.ZipFile(p + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            d = b.tighten(b.layout(d.decode("utf8"))).encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(p + ".tmp", p)
print("эталон: раскладка применена")
