#this program reads a xlsx file, returns the graph and table(optionally) in pdf format


import matplotlib.pyplot as plt
import pandas as pd
import dataframe_image as dfi
from fpdf import FPDF
import os
import cv2


# x = ""
def main():

    file1=File(fname(),ftype(), graph_title())
    file1.open_excel()

    file1.plot_points()
    file1.pdf_file()

    file1.delete_extra_files()

    print(file1.final_message())

#this function asks for file name from the user and returns the file name excluding the extension
def fname():
    fname=input("enter file name:").lower()
    if fname.endswith(".xlsx"):
        return fname[:-5]
    else:
        return fname


#this function asks for the file type from the user
def ftype():
    ftype=input("enter file type(formality ko lagi sodheko, i accept nothing but xlsx(excel)):").lower()

    if ftype==".xlsx" or ftype=="xlsx":
        return ".xlsx"
    else:
        raise TypeError("File needs to xlsx type")


def graph_title():
    graph_title=input("enter the title of your graph: ")    
    return graph_title


class File:

    def __init__(self,name,type,title):
        self.name=name
        self.type=type
        self.title=title

        self.fname=self.name+self.type

    #this function constructs data frame for first set of data
    def one_set(self):
        
        self.x_coordinate=80

        #stores the x and y label in respective variables
        self.x1_label=self.label_list[0]
        self.y1_label=self.label_list[1]

        #stores values of x and y in list
        self.x1_values=self.df[self.x1_label].tolist()
        self.y1_values=self.df[self.y1_label].tolist()

        if len(self.label_list)>2:
            self.two_set()
            return 0 

        self.data=[self.x1_values,self.y1_values]

        self.pandas_dataframe1()


    #this returns the csv values as pandas dataframe object when there is 1 set of data
    def pandas_dataframe1(self):
        self.dataframe=pd.DataFrame({self.x1_label:self.x1_values, self.y1_label:self.y1_values})
        self.table_title=self.title+"_table.png"
        dfi.export(self.dataframe,self.table_title)
        


    #this function constructs data frame for second set of data
    def two_set(self):
        
        self.x_coordinate=55

        #stores the x and y label in respective variables
        self.x2_label=self.label_list[2]
        self.y2_label=self.label_list[3]

        #stores values of x and y in list
        self.x2_values=self.df[self.x2_label].tolist()
        self.y2_values=self.df[self.y2_label].tolist()

        
        if len(self.label_list)>4:
            self.three_set()
            return 0

        self.data=[self.x1_values,self.y1_values,
                    self.x2_values,self.y2_values]
        self.pandas_dataframe2()


    #this returns the csv values as pandas dataframe object when there is 2 set of data
    def pandas_dataframe2(self):
        self.dataframe=pd.DataFrame({self.x1_label:self.x1_values, self.y1_label:self.y1_values,
                                self.x2_label:self.x2_values, self.y2_label:self.y2_values})
        self.table_title=self.title+"_table.png"
        dfi.export(self.dataframe,self.table_title)
        


    #this function constructs data frame for third set of data
    def three_set(self):
        
        self.x_coordinate=30

        #stores the x and y label in respective variables
        self.x3_label=self.label_list[4]
        self.y3_label=self.label_list[5]

        #stores values of x and y in list
        self.x3_values=self.df[self.x3_label].tolist()
        self.y3_values=self.df[self.y3_label].tolist()

        
        if len(self.label_list)>6:
            self.four_set()
            return 0

        self.data=[self.x1_values,self.y1_values,self.x2_values,
                    self.y2_values,self.x3_values,self.y3_values]
        self.pandas_dataframe3()


    #this returns the csv values as pandas dataframe object when there is 3 set of data
    def pandas_dataframe3(self):
        self.dataframe=pd.DataFrame({self.x1_label:self.x1_values, self.y1_label:self.y1_values,
                                self.x2_label:self.x2_values, self.y2_label:self.y2_values,
                                self.x3_label:self.x3_values, self.y3_label:self.y3_values})

        self.table_title=self.title+"_table.png"
        dfi.export(self.dataframe,self.table_title)
        

    
    #this function constructs data frame for fourth set of data
    def four_set(self):
        
        #stores the x and y label in respective variables
        self.x4_label=self.label_list[6]
        self.y4_label=self.label_list[7]

        #stores values of x and y in list
        self.x4_values=self.df[self.x4_label].tolist()
        self.y4_values=self.df[self.y4_label].tolist()

        
        if len(self.label_list)>8:
            self.five_set()
            return 0

        self.data=[self.x1_values,self.y1_values,
                    self.x2_values,self.y2_values,
                    self.x3_values,self.y3_values]
        self.pandas_dataframe4()


    #this returns the csv values as pandas dataframe object when there is 4 set of data
    def pandas_dataframe4(self):
        self.dataframe=pd.DataFrame({self.x1_label:self.x1_values, self.y1_label:self.y1_values,
                                self.x2_label:self.x2_values, self.y2_label:self.y2_values,
                                self.x3_label:self.x3_values, self.y3_label:self.y3_values,
                                self.x4_label:self.x4_values, self.y4_label:self.y4_values})

        self.table_title=self.title+"_table.png"
        dfi.export(self.dataframe,self.table_title)
        

    
    #this function constructs data frame for fifth set of data
    def five_set(self):
    
        #stores the x and y label in respective variables
        self.x5_label=self.label_list[8]
        self.y5_label=self.label_list[9]

        #stores values of x and y in list
        self.x5_values=self.df[self.x5_label].tolist()
        self.y5_values=self.df[self.y5_label].tolist()


        self.data=[self.x1_values,self.y1_values,
                 self.x2_values,self.y2_values,
                 self.x3_values,self.y3_values,
                 self.x4_values,self.y4_values,
                 self.x5_values,self.y5_values]
        self.pandas_dataframe5()


    #this returns the csv values as pandas dataframe object when there is 5 set of data
    def pandas_dataframe5(self):
        self.dataframe=pd.DataFrame({self.x1_label:self.x1_values, self.y1_label:self.y1_values,
                                self.x2_label:self.x2_values, self.y2_label:self.y2_values,
                                self.x3_label:self.x3_values, self.y3_label:self.y3_values,
                                self.x4_label:self.x4_values, self.y4_label:self.y4_values,
                                self.x5_label:self.x5_values, self.y5_label:self.y5_values})
        self.table_title=self.title+"_table.png"
        dfi.export(self.dataframe,self.table_title)
        

    #opens and stores values of x and y axis title and values

    def open_excel(self):

        #opening csv file
        self.df=pd.read_excel(self.fname)                  

        #label_list stores names of axes label as list
        self.label_list=self.df.columns.values.tolist()  

        self.one_set()


    def first_label(self):
        self.first=input("enter label for first set of data: ")
        return self.first


    def second_label(self):
        self.second=input("enter label for second set of data: ")
        return self.second
        

    def third_label(self):
        self.third=input("enter label for third set of data: ")
        return self.third

    
    def fourth_label(self):
        self.fourth=input("enter label for third set of data: ")
        return self.fourth

        
    def fifth_label(self):
        self.fifth=input("enter label for third set of data: ")
        return self.fifth

    #this method plots the graph with the points provided and sets title, places legend and saves the image file and shows the image
    def plot_points(self):
        plt.title(self.title)

        if len(self.label_list)==2:
            plt.plot(self.x1_values,self.y1_values,marker=".")

        if len(self.label_list)>2:
            plt.plot(self.x1_values,self.y1_values,marker=".",label=self.first_label())
            plt.plot(self.x2_values,self.y2_values,marker=".",label=self.second_label())

        if len(self.label_list)>4:
            plt.plot(self.x3_values,self.y3_values,marker=".",label=self.third_label())  

        if len(self.label_list)>6:  
            plt.plot(self.x4_values,self.y4_values,marker=".",label=self.fourth_label())

        if len(self.label_list)>8:
            plt.plot(self.x5_values,self.y5_values,marker=".",label=self.fifth_label())   

        if len(self.label_list)>2:
            plt.legend(loc="best")
             
        self.save_as=self.title+".png"
        plt.savefig(self.save_as)

        im=cv2.imread(self.save_as)
        self.table_width=im.shape[1]




    def length_of_table(self):
        return len(self.dataframe.index)

    def pdf_file(self):

        pdf=FPDF("P","mm","A4")

        pdf.add_page()
        pdf.set_margins(20,10,10)
        
        #adjusts the height of the table if it is just more than half a page
        if self.length_of_table()<15 and self.length_of_table()>10:
            #adjusts the width of the table is it is too wide
            if len(self.label_list)>6:
                pdf.image(self.table_title,w=180,h=130)
            else:
                pdf.image(self.table_title,x=self.x_coordinate,h=130)
        
            
        else:
            #adjusts the width of the table is it is too wide
            if len(self.label_list)>6:
                pdf.image(self.table_title,w=180)
            else:
                pdf.image(self.table_title,x=self.x_coordinate)
        pdf.image(self.save_as, w=170)

        pdf_save_as=self.title+".pdf"
        pdf.output(pdf_save_as)
        
    #this function deletes the images of the graph and table, resulting in only the required pdf file
    def delete_extra_files(self):
        if os.path.isfile(self.save_as):
            os.remove(self.save_as)
            
        if os.path.isfile(self.table_title):
            os.remove(self.table_title)

    #this function returns the message that pdf production is successful          
    def final_message(self):
        return ("\n\n\nthe pdf is ready!✨\n(search in the folder of this program📁, file name:title_of_graph.pdf), where title_of_graph is the title you entered!📰\n\n")
 
if __name__=="__main__":
    main()