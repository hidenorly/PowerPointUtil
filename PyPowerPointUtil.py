#   Copyright 2023 hidenorly
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.

import os
import string
import time

import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.dml.color import RGBColor
import webcolors


class PowerPointUtil:
    SLIDE_WIDTH_INCH = 16
    SLIDE_HEIGHT_INCH = 9

    def __init__(self, path):
        self.prs = Presentation()
        self.prs.slide_width  = Inches(self.SLIDE_WIDTH_INCH)
        self.prs.slide_height = Inches(self.SLIDE_HEIGHT_INCH)
        self.path = path

    def save(self):
        self.prs.save(self.path)

    # layout is full, left, right, top, bottom
    def getLayoutPosition(self, layout="full"):
        # for full
        x=0
        y=0
        width = self.prs.slide_width
        height = self.prs.slide_height

        if layout=="left" or layout=="right":
            width = width /2
        if layout=="top" or layout=="bottom":
            height = height /2
        if layout=="right":
            x=width
        if layout=="bottom":
            y=height

        return x,y,width,height

    def getLayoutToFitRegion(self, width, height, regionWidth, regionHeight):
        resultWidth = width
        resultHeight = height

        if width > height:
            resultWidth = regionWidth
            resultHeight = int(regionWidth * height / width+0.99)
        else:
            resultHeight = regionHeight
            resultWidth = int(regionHeight * width / height+0.99)

        return resultWidth, regionHeight


    def addSlide(self, layout=None):
        if layout == None:
            layout = self.prs.slide_layouts[6]
        self.currentSlide = self.prs.slides.add_slide(layout)

    def addPicture(self, imagePath, x=0, y=0, width=None, height=None, isFitToSlide=True, regionWidth=None, regionHeight=None, isFitWihthinRegion=False):
        if not regionWidth:
            regionWidth = self.prs.slide_width
        if not regionHeight:
            regionHeight = self.prs.slide_height
        regionWidth = int(regionWidth+0.99)
        regionHeight = int(regionHeight+0.99)
        pic = None
        try:
            pic = self.currentSlide.shapes.add_picture(imagePath, x, y)
        except:
            pass
        if pic:
            if width and height:
                pic.width = width
                pic.height = height
            else:
                if isFitToSlide:
                    width, height = pic.image.size
                    picWidth = pic.width
                    picHeight = pic.height
                    if width > height:
                        picWidth = regionWidth
                        picHeight = int(regionWidth * height / width + 0.99)
                    else:
                        picHeight = regionHeight
                        picWidth = int(regionHeight * width / height + 0.99)
                    if isFitWihthinRegion:
                        deltaWidth = picWidth - regionWidth
                        deltaHeight = picHeight - regionHeight
                        if deltaWidth>0 or deltaHeight>0:
                            # exceed the region
                            if deltaWidth > deltaHeight:
                                picWidth = regionWidth
                                picHeight = int(regionWidth * height / width + 0.99)
                            else:
                                picHeight = regionHeight
                                picWidth = int(regionHeight * width / height + 0.99)
                    pic.width = picWidth
                    pic.height = picHeight
        return pic

    def nameToRgb(name):
        result = RGBColor(0,0,0)
        try:
            rgb = webcolors.name_to_rgb(name)
            result = RGBColor(rgb.red, rgb.green, rgb.blue)
        except:
            pass
        return result

    def applyExFormat(exFormat, textbox, font, text_frame):
        exFormats = exFormat.split(",")
        for anFormat in exFormats:
            cmdarg = anFormat.split(":")
            cmd = cmdarg[0]
            val = None
            if len(cmdarg)>=2:
                val = cmdarg[1]
            if cmd=="color":
                font.color.rgb = PowerPointUtil.nameToRgb(val)
            elif cmd=="face":
                font.name = val
            elif cmd=="size":
                font.size = Pt(float(val))
            elif cmd=="bold":
                font.bold = True
            elif cmd=="effect":
                # TODO: fix
                shadow = textbox.shadow
                shadow.visible = True
                shadow.shadow_type = 'outer'
                shadow.style = 'outer'
                shadow.blur_radius = Pt(5)
                shadow.distance = Pt(2)
                shadow.angle = 45
                shadow.color = MSO_THEME_COLOR_INDEX.ACCENT_5
                shadow.transparency = 0

    def addText(self, text, x=Inches(0), y=Inches(0), width=None, height=None, fontFace='Calibri', fontSize=Pt(18), isAdjustSize=True, textAlign = PP_ALIGN.LEFT, isVerticalCenter=False, exFormat=None):
        if width==None:
            width=self.prs.slide_width
        if height==None:
            height=self.prs.slide_height
        width = int(width+0.99)
        height = int(height+0.99)

        textbox = self.currentSlide.shapes.add_textbox(x, y, width, height)
        text_frame = textbox.text_frame
        text_frame.text = text
        font = text_frame.paragraphs[0].font
        font.name = fontFace
        font.size = fontSize
        theHeight = textbox.height

        if exFormat:
            PowerPointUtil.applyExFormat(exFormat, textbox, font, text_frame)
        
        if isAdjustSize:
            text_frame.auto_size = True
            textbox.top = y

        if isVerticalCenter:
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        for paragraph in text_frame.paragraphs:
            paragraph.alignment = textAlign
