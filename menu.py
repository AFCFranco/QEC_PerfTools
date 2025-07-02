import generateExcelReport as gr
import jmxNamingConventions as anc
import compareReports as cr
from colorama import init, Fore, Style

print("""
════════════════════════════════════════════════════════════════
               QEC Performance Tools - v1.0 
════════════════════════════════════════════════════════════════""")
while True:
    print(Fore.RESET+Style.BRIGHT + "Select a task to execute:\n")
    print(Fore.CYAN + " 1.  Apply JMeter naming convention to a JMX file.")
    print(Fore.CYAN + " 2.  Generate Excel report.")
    print(Fore.CYAN + " 3.  Compare two Excel reports.")
    print(Fore.CYAN + " 4.  Exit\n")
    print()
    try:
        option = int(input(Fore.RESET+Style.DIM + "Type the option and press Enter: "))
        if option>4 or option<1:
            print("Invalid input.")
            break
        if option == 1:
            anc.applyJmxNamingConventions()
            break
        if option == 2:
            gr.genarateExcelreport()
            break
        if option == 3:
            cr.compareReports()
            break
        if option == 4:
            break
        continue
    except Exception:
        print("Invalid input.")

