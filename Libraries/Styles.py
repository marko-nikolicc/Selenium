from openpyxl.styles import Font, Alignment, colors

def stylePass():
    return Font(bold=True, color=colors.GREEN)

def styleFail():
    return Font(bold=True, color=colors.RED)

def styleBold():
    return Font(bold=True)

def position_center():
    return Alignment(horizontal="center")

def position_left():
    return Alignment(horizontal="left")