'''
#############################################

 PPT interface module
 Author: Tan Kok Hua (Guohua tan)
 Email: spider123@gmail.com

##############################################

'''

import win32com.client.dynamic 
import glob

class UsePPT(object): 
    '''
        Python PPT Interface. It provides methods for accessing the 
        basic functionality of MS PPT 97/2000 from Python. 

        This interface uses dynamic dispatch objects. All necessary constants 
        are embedded in the code. There is no need to run makepy.py. 
    ''' 

    def __init__(self, fileName=None):
        '''
            Initialize the ppt obj.
            str filename --> none
            if Filename == none, open a new presentation
        '''
        
        ## Initialize the parameters
        self.display_info = 0 #if 1- display the info mainly for debugging

        
        ## initialize the ppt application
        self.ppt_app = win32com.client.dynamic.Dispatch("PowerPoint.Application")
        
        if fileName:
            #need to activate to open a frame windows
            self.ppt_app.Activate()
            self.presentation = self.ppt_app.Presentations.Open(fileName)#open the frame window first 
        else: 
            self.presentation = self.ppt_app.Presentations.Add()
           
    def save(self, newfileName=None): 
        if newfileName:             
            self.presentation.SaveAs(newfileName) 
        else: 
            self.presentation.Save() 

    def close(self):
        '''
            Close the ppt obj
        '''
        self.presentation.Close() 
        del self.ppt_app 

    def hide(self): 
        self.ppt_app.Visible = False 

    def show(self): 
        self.ppt_app.Visible = True 

    ## Slide function
    def add_new_slide(self,page, page_layout =12):
        '''
            Add new slide to ppt
            12- default blank slide
        '''
        #12 -default blank slide
        self.presentation.Slides.Add(Index=page,Layout = page_layout)

    def delete_slide(self, page):
        self.presentation.Slides(page).Delete()

    def count_slide(self):
        return self.presentation.Slides.Count

    def num_of_shape_in_slide(self, slide_no):
        slide_obj = self.presentation.Slides(slide_no)
        return slide_obj.Shapes.Count

    ## !!!
    def move_to_target_slide(self,slide_no):
        '''
            Specify the sheet number to go to and activate that slide
            int (slide_no) --> None (activate the ppt)
        '''
    ## !!!
    def select_obj_for_particular_slide(self,slide_no):
        '''
            Shorten the abbreivation by returning the particular slide object
            

        '''
    
    ## Pic function
    def insert_pic_fr_file_to_slide(self,slide_no,pic_fname, xpos, ypos, size = (0,0) ):        
        '''
            Insert a pic from file to a slide specified by slide no
            int Slide no, str pic_fname, int xpos, int, ypos, tuple size (if size 0,0) will use the default size
            --> pic object (instance of pic object)
            tuple Size  (width, height)
            if size(0,0) will let the powerpoint decide the size
            Note:
                LinkToFile:=msoFalse (0),
                SaveWithDocument:=msoTrue (-1)
        '''
        if size  == (0,0):
            return self.presentation.Slides(slide_no).Shapes.AddPicture(FileName=pic_fname, LinkToFile=0, SaveWithDocument=-1, Left=xpos, Top=ypos)
        else:
            return self.presentation.Slides(slide_no).Shapes.AddPicture(FileName=pic_fname, LinkToFile=0, SaveWithDocument=-1, Left=xpos, Top=ypos,Width=size[0], Height=size[1])

    def paste_group_of_pic_according_to_num(self, slide_no, list_of_pic_fnames, layout = 'top'):
        '''
            from the list of fname determine the grouping required and paste the pictures accordingly.
            int slide_no, list of str list_of_pic_fname, str layout --> list of pic object
            
            NB: temporary have pasting for 1,2,3,4 (2 groups of 2),6 (2 groups of 3)
         
        '''
        if len(list_of_pic_fnames) == 2:
            list_of_pic_name = self.group_of_2_pic_pasting(slide_no, list_of_pic_fnames, layout)

        elif len(list_of_pic_fnames) == 4:
            list_of_pic_name1 = self.group_of_2_pic_pasting(slide_no, list_of_pic_fnames[:2], 'top')
            list_of_pic_name2 = self.group_of_2_pic_pasting(slide_no, list_of_pic_fnames[2:], 'bottom')
            list_of_pic_name = list_of_pic_name1 + list_of_pic_name2

        elif len(list_of_pic_fnames) == 3:
            list_of_pic_name = self.group_of_3_pic_pasting(slide_no, list_of_pic_fnames, layout, pic_to_pic_spacing_margin = -10)

        elif len(list_of_pic_fnames) == 6:
            list_of_pic_name1 = self.group_of_3_pic_pasting(slide_no, list_of_pic_fnames[:3], 'top', pic_to_pic_spacing_margin = -10)
            list_of_pic_name2 = self.group_of_3_pic_pasting(slide_no, list_of_pic_fnames[3:], 'bottom', pic_to_pic_spacing_margin = -10)
            list_of_pic_name = list_of_pic_name1 + list_of_pic_name2
            
        elif len(list_of_pic_fnames) ==1:
            ## arbitary first.
            pic_obj = self.insert_pic_fr_file_to_slide(slide_no,list_of_pic_fnames[0], 100, 100 )
            list_of_pic_name = [pic_obj.Name]
            
        else:
            print 'the len is not accuate. do nothing. '

        return list_of_pic_name


    ## !!!
    def group_pics_raw_pasting(self, slide_no, list_of_pic_fnames, start_pos_x, start_pos_y , spacing):
        '''
            Paste the pic in series of list (paste in horiztional sweeping)
            int slide no, list_of_pic_fname, int start pos_x, int start pos_y, spacing --> list_of_pic_name
            spacing -- distance betwee one plots to another

            Size of the plots will be default
        '''
        temp_xpos = start_pos_x
        list_of_pic_name = []
        for pic_fname in list_of_pic_fnames:
            pic_obj = self.insert_pic_fr_file_to_slide(slide_no,pic_fname, temp_xpos, start_pos_y, size = (0,0) )
            temp_xpos = temp_xpos + spacing
            list_of_pic_name.append(pic_obj.Name)

        return list_of_pic_name

    def group_of_2_pic_pasting(self, slide_no, list_of_pic_fnames, layout = 'top'):
        '''
            Handle pasting when the list of pic pass is 2.
            Paste in group of 2. (allow for multiple of two?
            int siide no, list list_of_pic_filename, layout {'top','bottom'} --> list_of_pic_name

            TODO: including squeeze factor? allow for 
            
        '''

        assert len(list_of_pic_fnames) == 2

        start_pos_x = 10
        spacing = 350
        
        if layout == 'top':
            start_pos_y = 30
        elif layout == 'bottom':
            start_pos_y = 200
        else:
            print 'layout is wrong, choose either top or bottom'
            raise

        return self.group_pics_raw_pasting(slide_no, list_of_pic_fnames, start_pos_x, start_pos_y , spacing)

    def group_of_3_pic_pasting(self, slide_no, list_of_pic_fnames, layout = 'top', further_margin = 0, pic_to_pic_spacing_margin = 0):
        '''
            To handle if the group of pic comes in 3
            int slide_no, list of str list_of_pic_fname, str layout ('top'/'bottom')
            int further_margin, int pic_to_pic_spacing_margin --> list of str (pic name)
            
            horizontal pasting -- either allow top or bottom pasting
            further_margin allow fine positioning of the top and botton position.
        '''
        assert len(list_of_pic_fnames) == 3

        list_of_pic_name = []

        if layout == 'top':
            if self.display_info: print 'top layout'           
            temp_xpos = 10
            ypos = 30 + further_margin
            for pic_fname in list_of_pic_fnames:
                pic_obj = self.insert_pic_fr_file_to_slide(slide_no,pic_fname, temp_xpos, ypos, size = (0,0) )
                temp_xpos = temp_xpos + 250 + pic_to_pic_spacing_margin
                list_of_pic_name.append(pic_obj.Name)
                
        elif layout == 'bottom':
            if self.display_info: print 'bottom layout'    
            temp_xpos = 10
            ypos = 200 + further_margin
            for pic_fname in list_of_pic_fnames:
                pic_obj = self.insert_pic_fr_file_to_slide(slide_no,pic_fname, temp_xpos, ypos, size = (0,0) )
                temp_xpos = temp_xpos + 250 + pic_to_pic_spacing_margin
                list_of_pic_name.append(pic_obj.Name)

        else:
            print 'Layout is not correct'
            raise

        return list_of_pic_name


    def select_obj_of_series_of_pic(self,slide_no, list_of_pic_name):
        '''
            Select a group of pic for a slide no based on the pic name
            and return the object so that it can be use for other sequential function
            int slide_no, list of str (list_of_pic_name) --> shape range object

        '''
        return self.presentation.Slides(slide_no).Shapes.Range(list_of_pic_name)

    def scale_selected_grp_of_pic(self, picrange_obj, scalingfactor = 0.8, group_select =1):
        '''
            Take in a series of selected object and perform scaling according to scaling factor (1 is original size).
            this is uniform scaling with width and height scale equally.

            (pic range object from select_obj_of_series_of_pic), scaling factor, boolean group_select --> None

            group_select: group the object before scaling.
            After Scaling, revert back to ungroup to preserve each individual picture name
            
            PPT VBA algo:
                .ScaleWidth 0.83, msoFalse (0), msoScaleFromTopLeft (0)
                .ScaleHeight 0.83, msoFalse, msoScaleFromTopLeft

        '''
        if group_select:
            picrange_obj = picrange_obj.Group()

        picrange_obj.ScaleWidth(scalingfactor,0,0)
        picrange_obj.ScaleHeight(scalingfactor,0,0)

        if group_select:
            picrange_obj.Ungroup() # to release the grouping to prevent any naming change

    def adjust_pos_selected_grp_of_pic(self, picrange_obj, direction = 'vertical', offset = [0,0], group_select =1):
        '''
            Take in a series of selected object and perform vertical and horizontal adjustment according to offset value

            (pic range object from select_obj_of_series_of_pic), str direction (vertical/horizontal/both),
            list of two offset, boolean group_select --> None
            offset[vertical, horizontal]
            group_select: group the object before scaling.
            offset, go negative means go up
            After adjust, revert back to ungroup to preserve each individual picture name

        '''

        if group_select:
            picrange_obj = picrange_obj.Group()

        if direction == 'vertical':
            picrange_obj.Top = picrange_obj.Top + offset[0]
        elif direction == 'horizontal':
            picrange_obj.Left = picrange_obj.Left + offset[1]
        elif direction == 'both':
            picrange_obj.Top = picrange_obj.Top + offset[0]
            picrange_obj.Left = picrange_obj.Left + offset[1]
        else:
            print 'Direction not valid. No action taken'

        if group_select:
            picrange_obj.Ungroup() # to release the grouping to prevent any naming change


    def crop_selected_grp_of_pic(self, picrange_obj, crop_amt, crop_type = 'bottom'):
        '''
            Function to crop a series of pic to required size
            picrange_obj, string crop_type (bottom, right), float crop_amt --> None

            Excel VBA code
            ActiveWindow.Selection.ShapeRange.PictureFormat.CropBottom = 196.96
        '''
        if crop_type == 'bottom':
            picrange_obj.PictureFormat.CropBottom = crop_amt
        elif crop_type == 'right':
            picrange_obj.PictureFormat.CropRight = crop_amt
        else:
            print 'Crop Type not found, choose bottom or right'
            raise


    ##!!!
    def align_select_grp_of_pic(self,picrange_obj, align_type = 'center'):
        '''
            Take in a series of selected object and perform aligning according to alignment_type
            picrange_obj, str alignment_type --> none

            ActiveWindow.Selection.ShapeRange.Align msoAlignTops, (3) False

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

        alignment_dict = {
                            'left': 0,
                            'center': 1,
                            'right': 2,
                            'top': 3,
                            'middle': 4,
                            'bottom': 5

                            }

        if not align_type in algnment_dict.keys():
            print 'alignment type not recognised'
            print 'no effect on alignment'

        picrange_obj.Align(alignment_dict[align_type], False)

    ## Text box function
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

    def get_tar_textbox(self, slide_no, search_txt):
        '''Get target textbook based on search text?'''
        shape_no_list = self.get_shapes_of_tar_type(slide_no,type = 17,display_raw=0)
        for n in shape_no_list:
            shape_obj = self.presentation.Slides(slide_no).Shapes(n)
            temp_text = shape_obj.TextFrame.TextRange.Text
            if temp_text.find(search_txt)>=0:
                return n
        return None


    ## Info functon
    def display_shape_obj_properties(self, slide_no, shape_no):
        '''
            textbox =17,pic = 13, placeholder =14
            return in ?dict
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

    def get_all_shapes_properties(self, slide_no, print_output =0):
        """Module to display and return all shapes properties.
            Args:
                slide_no (int):

            Returns:
                (list): return the list of all the shape properties in  'Name,Left,Top, Ht, Width, Type'

            
        """

        all_shape_properties_list = list()
        for n in range(self.num_of_shape_in_slide(slide_no)):
            all_shape_properties_list.append(self.display_shape_obj_properties(slide_no,n+1))

        if print_output:
            print 'Name,Left,Top, Ht, Width, Type'
            for n in all_shape_properties_list: print n
            
        return all_shape_properties_list

    ## Shapes and position adjustment
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
                
                
if (__name__ == "__main__"):
    #for testing purpose
    test = 1

    if test ==1:    
        ppt_app = win32com.client.dynamic.Dispatch("PowerPoint.Application")
        filename = r'C:\Documents and Settings\Tan Kok Hua\Desktop\summary.ppt'
        filename = r'C:\data\summary.ppt'
        ppt_app.Activate()
        #presentation = ppt_app.Presentations.Open(filename)
        presentation = ppt_app.Presentations.Add()
        ppt_app.Visible=1
        #presentation.Close()
        #ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutBlank).SlideIndex -->12
        #presentation.Slides.Add(index=2,Layout =12)--> may be abloe to return the current object
        #delete slides --> presentation.Slides(1).Delete()
        #slides count


        
   


