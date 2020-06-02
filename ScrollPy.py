import subprocess
import tkinter as tk
from tkinter import StringVar
from tkinter import filedialog
from tkinter import ttk
import xlrd
from PIL import Image, ImageDraw, ImageFont, ImageTk

HEIGHT = 1920
WIDTH = 1080
W, H = 0, 0

fps = 23.98
scrollspeed = 2
ccc_file = " "
fontsize = 16
linecount = 0
pngsave = "C:\\Change\\this\\Directory\\credits.png"
png = "credits.png"


def get_excel():
    global ccc_file
    ccc_file = filedialog.askopenfilename(initialdir="/", title="Select xml file",
                                          filetypes=(("spreadsheet files", "*.xls"), ("all files", "*.*")))
    entry.configure(state='normal')
    entry.delete(0, 'end')
    entry.insert(0, ccc_file)
    entry.configure(state='readonly')


def draw_underlined_text(draw, pos, text, font, **options):
    twidth, theight = draw.textsize(text, font=font)
    lx, ly = pos[0], pos[1] + (theight*1.1)
    draw.text(pos, text, font=font, **options)
    draw.line((lx, ly, lx + twidth, ly), **options)


def set_fps():
    fps = fps_combo.get()


def CreateScroll(video, output):
    global linecount
    global progress
    progress.start()
    fps = fps_combo.get()
    scrollspeed = scroll_widget.get()
    command = f'ffmpeg -y -f lavfi -i "color=black:s="{W_entry.get()}"x"{H_entry.get()}", fps=fps="{fps}"[background]; movie="{video}""[overlay];[background][overlay]overlay="0:H-(n*"{scrollspeed}")",smartblur=0.75"" -vframes "{(int(H) / int(scrollspeed)) + (int(H_entry.get()) / int(scrollspeed))}" -framerate "{fps}" -preset veryfast -crf 18 -v verbose ""{output}""'
    subprocess.call(command, stdout=None, stderr=None, shell=True,
                    cwd="C:/Users/ambus/Desktop/Scripting/ScrollingCreds")
    command2 = f'ffplay "{output}" -x 500'
    subprocess.Popopen(command2, stdout=None, stderr=None, shell=True,
                    cwd="C:/Users/ambus/Desktop/Scripting/ScrollingCreds")
    progress.stop()

def generate_preview():
    workbook = xlrd.open_workbook(ccc_file)
    sheet = workbook.sheet_by_name("Scroll")
    leftvalues = sheet.col_values(1)
    leftvalues = [" " if x == '' else x for x in leftvalues]
    leftvalues = leftvalues[1:]
    rightvalues = sheet.col_values(3)
    rightvalues = [" " if x == '' else x for x in rightvalues]
    rightvalues = rightvalues[1:]
    centervalues = sheet.col_values(2)
    centervalues = [" " if x == '' else x for x in centervalues]
    centervalues = centervalues[1:]
    cardvalues = sheet.col_values(5)
    cardvalues = [" " if x == '' else x for x in centervalues]
    cardvalues = centervalues[1:]
    leftcount = len(leftvalues)
    centercount = len(centervalues)
    rightcount = len(rightvalues)

    lines_var = int(max(leftcount, centercount, rightcount))
    print(lines_var)
    # draw with PIL
    global W, H
    W, H = (int(W_entry.get()), 20 * lines_var)
    print(H)
    im = Image.new("RGBA", (W, H), "black")
    fontsize = int(W / 100)
    myFont = ImageFont.truetype("constan.ttf", fontsize)
    myFontbold = ImageFont.truetype("constan.ttf", fontsize)
    gutterSize = 25
    # leftCreds
    LineCounter = 0
    for x in leftvalues:
        LineCounter = LineCounter + 1
        # print(x)
        msg = str(x)
        draw = ImageDraw.Draw(im)
        w, h = myFont.getsize(msg)
        draw.text(((W / 2) - gutterSize - w, 0 + (fontsize * LineCounter)), msg, font=myFont, fill="white")
    leftcount = LineCounter
    # centerCreds
    LineCounter = 0
    for x in centervalues:
        LineCounter = LineCounter + 1
        # print(x)
        msg = str(x)
        draw = ImageDraw.Draw(im)
        w, h = myFont.getsize(msg)
        if x == "CAST" or x == "ALBERTA CREW" or x == "Special Thanks to":
            # draw.text(((W/2)-(w/2),0+(fontsize*LineCounter)), msg, font=myFont, fill="white")
            draw_underlined_text(draw, ((W / 2) - (w / 2), 0 + (fontsize * LineCounter)), msg, font=myFont,
                                 fill="white")
        else:
            draw.text(((W / 2) - (w / 2), 0 + (fontsize * LineCounter)), msg, font=myFont, fill="white")
    centercount = LineCounter
    # print(centercount)
    # rightCreds
    LineCounter = 0
    for x in rightvalues:
        LineCounter = LineCounter + 1
        # print(x)
        msg = str(x)
        draw = ImageDraw.Draw(im)
        w, h = myFont.getsize(msg)
        draw.text(((W / 2) + gutterSize, 0 + (fontsize * LineCounter)), msg, font=myFont, fill="white")
    rightcount = LineCounter
    # print(leftcount)
    # print(centercount)
    # print(rightcount)
    global linecount
    linecount = int((int(H_entry.get()) / int(lines_var) * 1000) - (scrollspeed * 1000))
    # print(linecount)

    im.save(pngsave, "PNG")

    #myFont = ImageFont.truetype("my-font.ttf", 16)
    #draw.textsize(msg, font=myFont)
    # time.sleep(10)
    # calculate end time by number of lines * H
    global filename
    img = Image.open(filename)
    img2 = img.copy()
    Image.Image.close(img)
    resized_img = img2.resize( [int(.5 * s) for s in img2.size] )
    photoimg = ImageTk.PhotoImage(resized_img)
    labelimage.configure(image=photoimg)
    labelimage.image = photoimg


def generate_movie():
    global linecount
    #progress.start()
    #Working FFMPEG Command ffmpeg -f lavfi -i "color=black:s=1920x1080, fps=fps=23.98[background]; movie=credits.png[overlay];[background][overlay]overlay='0:H-(n*2)',smartblur=0.75" -t (3000/23.98) -framerate 23.976 -preset veryfast -crf 18 output2.mp4
    CreateScroll(png, filedialog.asksaveasfilename(initialdir="/", filetypes=(("mp4 files", "*.mp4"), ("all files", "*.*"))))

root = tk.Tk()
root.title("CreditCreator")
# create a toplevel menu
menubar = tk.Menu(root)

# create a pulldown menu, and add it to the menu bar

filemenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="Open xls", command=get_excel)
filemenu.add_command(label="Save", command="")
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)

menubar.add_command(label="Quit!", command=root.quit)

# display the menu
root.config(menu=menubar)

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

Width_var = StringVar()
Height_var = StringVar()

frame = tk.Frame(root, bg='#A9A9A9')
frame.place(relx=0, rely=0, relwidth=0.5, relheight=1)

frameout = tk.Frame(root, bg='gray')
frameout.place(relx=.5, rely=0, relwidth=0.5, relheight=1)

main_label = tk.Label(frame, text="Credit Creator", anchor="center", font=("Courier",44))
main_label.place(relx=0, rely=.1, relwidth=1)

button = tk.Button(frame, text="Choose Excel", bg='gray', command=get_excel)
button.place(relx=0, rely=.2, relwidth=.3, relheight=.03)

buttonPrev = tk.Button(frame, text="preview", bg='gray', command=generate_preview)
buttonPrev.place(relx=.5 - (.3 / 2), rely=.6, relwidth=.3, relheight=.05)

buttonrun = tk.Button(frame, text="Run", bg='gray', command=generate_movie)
buttonrun.place(relx=.5 - (.3 / 2), rely=.7, relwidth=.3, relheight=.05)

fps_combo = ttk.Combobox(frame, textvariable=fps, state='readonly')
fps_combo['values'] = [23.98, 24, 25, 29.97, 30, 48, 60]
fps_combo.current(0)
fps_combo.place(relx=0, rely=.27, relwidth=.3, relheight=.03)

scroll_widget = tk.Scale(frame, from_=0, to=10, orient=tk.HORIZONTAL)
scroll_widget.set(2)
scroll_widget.place(relx=.3, rely=.35, relwidth=.7, relheight=.04)

scrolllabel = tk.Label(frame, text="scroll speed")
scrolllabel.place(relx=0, rely=.35, relwidth=.3, relheight=.04)

label = tk.Label(frame, text="choose raster")
label.place(relx=.5 - (.2 / 2), rely=.45, relwidth=.2, relheight=.03)

entry = tk.Entry(frame, bg='white', state='readonly')
entry.place(relx=0.3, rely=.2, relwidth=.7, relheight=.03)

W_entry = tk.Entry(frame, bg='white', textvariable=Width_var)
W_entry.place(relx=0.3, rely=.5, relwidth=.2, relheight=.02)
W_entry.insert(0, "1920")

H_entry = tk.Entry(frame, bg='white', textvariable=Height_var)
H_entry.place(relx=0.5, rely=.5, relwidth=.2, relheight=.02)
H_entry.insert(0, "1080")

progress = ttk.Progressbar(root, mode ="indeterminate")
progress.place(anchor='center', rely=.9, relwidth=1, relheight=.03)

filename = 'C:\\Change\\this\\Directory\\credits.png'
img = Image.open(filename)
img2 = img.copy()
Image.Image.close(img)

resized_img = img2.resize( [int(.5 * s) for s in img2.size] )

# root = tk.Tk()
photoimg = ImageTk.PhotoImage(resized_img)
labelimage = tk.Label(frameout, image=photoimg)
labelimage.place(relx=0, rely=0, relwidth=1)

root.mainloop()
