import os, shutil

NAS   = r"Z:\종소세2026\고객"
LOCAL = r"C:\Users\pc\종소세2026"
os.makedirs(LOCAL, exist_ok=True)

names = ["강유진","김성준","김지은","나기은","박현민","변은지",
         "오상연","이근만","이명회","이선웅","정재호","지성환"]

for name in names:
    folder = os.path.join(NAS, name)
    fn  = f"{name}_종합소득세.pdf"
    src = os.path.join(folder, fn)
    dst = os.path.join(LOCAL, fn)
    if os.path.exists(src):
        shutil.move(src, dst)
        print(f"복구: {fn}")
    if os.path.isdir(folder) and not os.listdir(folder):
        os.rmdir(folder)
        print(f"폴더 삭제: {name}")

print("완료")
