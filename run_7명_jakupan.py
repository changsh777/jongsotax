"""
run_7명_jakupan.py — 7명 작업판 생성
이혜주/이완호/배성섭/박형우/김지혁/이재윤/두봉수

실행: python F:\종소세2026\run_7명_jakupan.py
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from jakupan_gen import make_jakupan

TARGETS = [
    ("이혜주", "900307"),
    ("이완호", "701228"),
    ("배성섭", "700510"),
    ("박형우", "950630"),
    ("김지혁", "780513"),
    ("이재윤", "970502"),
    ("두봉수", "690204"),
]

def main():
    print(f"=== 7명 작업판 생성 시작 ===\n")
    ok, fail = [], []
    for name, jumin6 in TARGETS:
        print(f"[{name}] 생성 중...")
        try:
            result = make_jakupan(name, jumin6)
            if result:
                print(f"  ✅ 완료: {result}")
                ok.append(name)
            else:
                print(f"  ⚠️  생성 실패 (None 반환)")
                fail.append(name)
        except Exception as e:
            print(f"  ❌ 오류: {e}")
            fail.append(name)

    print(f"\n=== 완료: 성공 {len(ok)}명 / 실패 {len(fail)}명 ===")
    if ok:   print("  성공:", ", ".join(ok))
    if fail: print("  실패:", ", ".join(fail))

if __name__ == "__main__":
    main()
