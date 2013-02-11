'''
#############################################

 PPT interface module
 Author: Tan Kok Hua (Guohua tan)
 Email: kokhua81@gmail.com
 Revised date: Feb 11 2013

 Executing simple PPT commands via python
##############################################

'''

import win32com.client.dynamic 
import glob

class UsePPT(object): 
    '''
        Python PPT Interface. It provides methods for accessing the 
        basic functionality of MS PPT 97/2000 from Python.
        
    ''' 

    def __init__(self, fileName=None):
        '''
            Open a new ppt or existing ppt.show

            Open a new ppt if fileName is not specified.
    
        '''
        
        self.ppt_app = win32com.client.dynamic.Dispatch("PowerPoint.Application")
        
        if fileName:
            self.ppt_app.Activate()#need to activate to open a frame windows
            self.presentation = self.ppt_app.Presentations.Open(fileName)#open the frame window first 
        else: 
            self.presentation = self.ppt_app.Presentations.Add()
           
    def save(self, newfileName=None): 
        if newfileName:             
            self.presentation.SaveAs(newfileName) 
        else: 
            self.presentation.Save() 

    def close(self):
        self.presentation.Close() 
        del self.ppt_app 

    def hide(self): 
        self.ppt_app.Visible = False 

    def show(self): 
        self.ppt_app.Visible = True 

    def add_new_slide(self,page, page_layout =12):
        '''
            Add new slide to ppt
            page_layout = 12-> default blank slide
        '''
        self.presentation.Slides.Add(index=page,Layout = page_layout)

    def delete_slide(self, page):
        self.presentation.Slides(page).Delete()

    def count_slide(self):
        return self.presentation.Slides.Count
        
    def add_textbox(self,page, pos_list, text):
        '''
            Adding a textbox of pos indicated by pos_listb(list type)--> start x, start y, end x, end y
            In the function, the text orientation will be added as part of args (it is default to be horizontal)
            ?To have a check on the textbox dimension

            Return tx_box obj            
            
        '''

        pos_list.insert(0,1)#1- msoTextOrientationHorizontal
        tb_obj = self.presentation.Slides(page).Shapes.AddTextbox(pos_list[0],pos_list[1],pos_list[2],pos_list[3],
                                                                  pos_list[4])
        tb_obj.TextFrame.TextRange.Text = text
        return tb_obj

    def get_txtbox_text(self,page,shape_index):
        '''
            Check for exception to make sure target selected is text box
        '''
        shape_obj = self.presentation.Slides(page).Shapes(shape_index)
        if not shape_obj.Type == 17:
            print 'Object not a textbox'
            return
        return shape_obj.TextFrame.TextRange.Text

    def change_textbox_text_style(self, page, shape_index, chr_start, chr_end, font_size = 24, font_bold =0,font_italics= 0):
        '''
            if chr_start = -1, change the whole sentence (not sure this is workiing?)
            underconstruction: font size, color, bold , italics
            ? can take in obj or just page
        '''
        shape_obj = self.presentation.Slides(page).Shapes(shape_index)
        tar_chr_select = shape_obj.TextFrame.TextRange.Characters(chr_start,chr_end)
        tar_chr_select.Font.Bold = font_size
        tar_chr_select.Font.Bold = font_bold
        tar_chr_select.Font.Italic = font_italics 

    def num_of_shape_in_slide(self, slide_no):
        slide_obj = self.presentation.Slides(slide_no)
        return slide_obj.Shapes.Count

    def display_shape_obj_properties(self, slide_no, shape_no):
        '''
            textbox =17,pic = 13, placeholder =14
            return in dict
            return name, pos left, pos top, ht, width, obj type
        '''
        shape_obj = self.presentation.Slides(slide_no).Shapes(shape_no)
        return shape_obj.Name, shape_obj.Left, shape_obj.Top, shape_obj.Height, shape_obj.Width, shape_obj.Type

    def get_shapes_of_tar_type(self, slide_no,type = 17,display_raw=1):
        '''
            textbox =17,pic = 13, placeholder =14
            return the shape no list of particular type
            underconstruciton: return obj, have args that serarch for name filter
        '''
        shape_no_list = list()
        for n in range(self.num_of_shape_in_slide(slide_no)):
            shape_obj = self.presentation.Slides(slide_no).Shapes(n+1)
            if shape_obj.Type == type:
                if display_raw: print shape_obj.Name, shape_obj.Left, shape_obj.Top, shape_obj.Height, shape_obj.Width, shape_obj.Type
                shape_no_list.append(n+1)
        return shape_no_list

    def get_tar_textbox(self, slide_no, search_txt):
        shape_no_list = self.get_shapes_of_tar_type(slide_no,type = 17,display_raw=0)
        for n in shape_no_list:
            shape_obj = self.presentation.Slides(slide_no).Shapes(n)
            temp_text = shape_obj.TextFrame.TextRange.Text
            if temp_text.find(search_txt)>=0:
                return n
        return None
        
    def display_all_shapes_properties(self, slide_no):
        print 'Name,Left,Top, Ht, Width, Type'
        for n in range(self.num_of_shape_in_slide(slide_no)):
            print self.display_shape_obj_properties(slide_no,n+1)

    def adjust_shape_pos(self, slide_no, shape_no, left,top, height, width):
        shape_obj = self.presentation.Slides(slide_no).Shapes(shape_no)
        shape_obj.Left = left
        shape_obj.Top = top
        shape_obj.Height = height
        shape_obj.Width = width 

    def align_shapes(self,slide_no, shape_no, shape_index_array, alignment):
        '''
            Alignment will be 1 t0 4: Use dict to switch
            align can be string or int
            left 0
            center 1
            right 2
            top 3
            middle 4
            bottom 5
            align top = 3
        '''
        shape_obj_group = self.presentation.Slides(slide_no).Range.Shapes(shape_index_array)
        shape_obj_group.Align(alignment,False)

    def configure_pic_pos(no_of_pic, remarks):
        '''
            take in no of pic and calculate its relative pos
            format type separate or together
            return top and left pos
            remarks -- top,middle,bottom ; spread, concentrate
            remarks -- [list - position -y diretion, list - action, placement - vertical horizontal]
            
        '''
        total_width = 720
        total_ht = 540
        obj_width =300
        obj_ht =200

        #instead of procedure, str away give obj position in array
        if no_of_pics == 2:
            #output will be a dict
            positioning_dict = {'top':100,'bottom':  300}
            if remarks == '':
                pass
                
                
def run_custom_test():
    #add other test here
    ppt = UsePPT() 
    ppt.show()
    ppt.add_new_slide(1)
    return ppt
    
                


if (__name__ == "__main__"):
    #for testing purpose
    test = 1

    if test == 1:
        ppt = run_custom_test()


    if test == 2:    
        ppt_app = win32com.client.dynamic.Dispatch("PowerPoint.Application")
        filename = r'C:\Documents and Settings\Tan Kok Hua\Desktop\summary.ppt'
        filename = r'C:\data\summary.ppt'
        ppt_app.Activate()
        presentation = ppt_app.Presentations.Open(filename)
        #presentation = ppt_app.Presentations.Add()
        ppt_app.Visible=1
        #presentation.Close()
        #ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutBlank).SlideIndex -->12
        #presentation.Slides.Add(index=2,Layout =12)--> may be abloe to return the current object
        #delete slides --> presentation.Slides(1).Delete()
        #slides count


        
    if test ==2:
        filename = r'C:\data\summary.ppt'
        ppt = UsePPT(filename) 
        ppt.show()