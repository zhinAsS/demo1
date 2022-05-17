from time import sleep
import xlwt
from  selenium import webdriver
# 由于时间原因，没有用框架实现，也还有很多需要优化的地方。

class Demo():
    def __init__(self):
#获取驱动，打开浏览器
        wb=webdriver.Chrome()
        wb.get('https://opensea.io/collection/catalog-lu-store')
        wb.maximize_window()
        wb.implicitly_wait(5)
        self.wb=wb
        # 计行，创建Excel
        self.r=0
        self.excel=xlwt.Workbook(encoding='utf8')
        self.sheet=self.excel.add_sheet('rest',cell_overwrite_ok=True)


    def demo(self):
        #首次向下滚动距离
        y=700
        js=f"window.scrollTo(0,{y});"
        num_item=1 #对写入的文件计数
        while num_item<2490:
            lis=[]
            self.wb.execute_script(js)
            # 获取当前页面项目的个数
            sum_item=self.wb.find_elements('xpath','//div[@class="Blockreact__Block-sc-1xf18x6-0 Assetreact__StyledContainer-sc-bnjqwy-0 elqhCm bwCDxg Asset--loaded"]/article/a')
            # 对点过的item计数
            # print('num_item>>',num_item)
            for i in range(1,len(sum_item)+1):
                # 找到项目，获取url地址
                url= self.wb.find_element('xpath',f'(//div[@class="Blockreact__Block-sc-1xf18x6-0 Assetreact__StyledContainer-sc-bnjqwy-0 elqhCm bwCDxg Asset--loaded"]/article/a)[{i}]').get_attribute('href')
                # 点击
                button=self.wb.find_element('xpath',f'(//div[@class="Blockreact__Block-sc-1xf18x6-0 Assetreact__StyledContainer-sc-bnjqwy-0 elqhCm bwCDxg Asset--loaded"]/article/a)[{i}]')
                self.wb.execute_script("(arguments[0]).click()",button)
            # 点击刷新按钮
                self.wb.find_element('xpath','//i[@value="refresh"]').click()
            # 获取提示弹框文本
                sleep(2)
                text=self.wb.find_element('xpath','//div[@class="Toastreact__DivContainer-sc-6g7ouf-0 fASrMR"]/div').text
                if text.split('\n')[1]=="We've queued this item for an update! Check back in a minute...":
                    status='Queued'
                else:
                    status='Error'

            # 定义一个临时列表用来放一个项目的数据：lis,
                temp_lis=[]
                temp_lis.append(num_item)
                temp_lis.append(url)
                temp_lis.append(status)
                lis.append(temp_lis)
                print('lis>>>>',lis)
                num_item+=1
                # 后退
                self.wb.back()
                sleep(2)
            self.writer(lis)
            js=f"window.scrollTo(0,{500});"

    # 写入Excel数据
    def writer(self,lis):
            for row in range(len(lis)):
                for col in range(len(lis[row])):
                    self.sheet.write(self.r,col,lis[row][col])
                    print(self.r,col)
                self.r+=1
            self.excel.save('./rest.xls')
if __name__ == '__main__':
    Demo().demo()