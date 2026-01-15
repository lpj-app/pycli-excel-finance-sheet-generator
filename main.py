import sys
from textConfig import TEXTS
from sheetGenerator import create_excel
from datetime import datetime

current_date = datetime.now()
current_year = current_date.year

def get_input(prompt, valid_options):
    while True:
        val = input(prompt).lower().strip()
        if val in valid_options:
            return val
        print(f"Invalid input. Please choose: {', '.join(valid_options)}")

def main():
    try:
        print("=========================================")
        print("     FINANCIAL SHEET GENERATOR v1.0   ")
        print("  Create your personal budgeting Excel  ")
        print("                                         ")
        print("© Luca-Pascal Junge - github.com/lpj.app  ")
        print("=========================================")
        
        # 1. Choose language
        print("\n[1] German")
        print("[2] English")
        lang_choice = get_input(">> Choose Language (1/2): ", ["1", "2"])
        
        lang_code = "de" if lang_choice == "1" else "en"
        current_text = TEXTS[lang_code]
        
        print(f"\nLanguage selected: {lang_code.upper()}")

        # 2. Choose year
        year = input(f">> Enter Year (Default {current_year}): ").strip()
        if not year: year = str(current_year)

        # 3. Choose mode
        print("\n[1] Single Sheet (Only the planner)")
        print("[2] Full Dashboard (Start page + Charts + Manual)")
        mode_choice = get_input(">> Select Mode (1/2): ", ["1", "2"])
        
        is_dashboard = (mode_choice == "2")
        mode_text = "Dashboard Template" if is_dashboard else "Single Sheet"
        
        # Generating
        print(f"\nGenerating {mode_text} for {year} in {lang_code.upper()}...")
        
        try:
            filename = create_excel(year, lang_code, current_text, with_dashboard=is_dashboard)
            print(f"\nSUCCESS! File created: {filename}")
            print("   You can now open it in Excel.")
        except Exception as e:
            print(f"\nERROR: {e}")
            import traceback
            traceback.print_exc()

        input("\nPress Enter to exit...")

    except KeyboardInterrupt:
        # Handles Ctrl+C gracefully
        print("\n\nAborted by user. Exiting...")
        sys.exit(0)

if __name__ == "__main__":
    main()