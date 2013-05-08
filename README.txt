PDFtoPPTX 0.2
############

Motivation
##################################################################################
The essential idea of this project evolved from my wish to use the Power Point
presenter view for my LaTex created slides that are in PDF. So it’s actually no
fancy wish I thought but trying to use the Adobe implemented options resulted in
quite crappy output. Since I didn’t want to pay for any third party software
that didn’t offer previews of their output I decided to code something myself.
This how my first project in C# came to life.

How it works
##################################################################################
What you find in this repository is basically the second version of my windows
from application. The small tool basically takes a snapshot of each slide of the
PDF presentation and imports it into a new Power Point presentation. You need
Power Point and a the Adobe Acrobat Pro/Reader to use the program. The user
specifies the location of a dummy.potx that is used as target for the snapshots,
the path of the source PDF, the number of slides, the output path and some place
to store the temporary snapshots. The quality of the resulting pptx is great and
after “conversion” you can add comments in Power Point that are available in the
presenter mode in Power Point. The minor drawback is, that any links or buttons
included in the original PDF are not preserved since the slides are copied as
pictures.

What I did and didn’t do
##################################################################################
So far I used Visual Studio and C# to write the program. All actions performed
on the PDF are done via sendKeys commands. I know it’s not the best option but
it’s my first C program ever and I thought that will do. I couldn’t get a direct
Adobe access to work, so if anyone has a better solution it will be much
appreciated. For the Power Point operations I used the implemented Office
libraries in Visual Studio to command Power Point “directly”. In the first
versions this was also done via sendKeys. It has some controls implemented to
check if all necessary inputs are made and returns an error message if not. I
didn’t find a way around the Office protection against downloaded files, so maybe
one has to manually approve the dummy.potx, open it confirm the dialog, or just
create an empty PPTX as target. Another small issue is that it’s not working in
the background. So as long as the conversion runs you cannot make any input, but
since it’s quite fast, on an average PC, I didn’t bother to work on that.

At this time it is certainly not perfect but it’s serves the purpose it was
created for. I find it quite useful therefore wanted to share it. If somebody
knows how to improve it you are welcome to it.

So fare it works with of Power Point 07/10/12 and Adobe Acrobat/Reader 11 on
Windows 7. The short cuts for Adobe Acrobat/Reader 9 are in the comments of the
source code. I couldn’t get a version selection to work so I just left it in the
comments.

Repository Content
##################################################################################
This repository includes the Visual Studios Project Folder, the release folder
including the dummy.potx and the actual program .exe file, and an additional file
with source code in C#.
