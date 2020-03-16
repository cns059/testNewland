package com.utils;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Iterator;

/**
 * @author chenningshui
 * @date 2020/3/13 9:10
 */
public class ExcelUtil {
    @SuppressWarnings("unchecked")
    // 创建excel文件函数
    // src为待保存的文件路径,json为待保存的json数据
    //data为json数组的数组名
    public static JSONObject createExcel(String src, JSONObject json,String data) {
        JSONObject result = new JSONObject(); // 用来反馈函数调用结果
        try {
            // 新建文件
            File file = new File(src);
            file.createNewFile();
            OutputStream outputStream = new FileOutputStream(file);// 创建工作薄
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(outputStream);
            WritableSheet sheet = writableWorkbook.createSheet("First sheet", 0);// 创建新的一页
            JSONArray jsonArray = json.getJSONArray(data);// 得到data对应的JSONArray
            Label label; // 单元格对象
//            int column = 0; // 列数计数\
            int rownum=0;//行级计数
            // 将第一行信息加到页中。如：姓名、年龄、性别
            label = new Label(0,0, "名称"); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值
            sheet.addCell(label); // 将单元格加到页
            label = new Label(1,0, "值"); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值
            sheet.addCell(label); // 将单元格加到页
            JSONObject first = jsonArray.getJSONObject(0);
            Iterator<String> iterator = first.keys(); // 得到第一项的key集合
            while (iterator.hasNext()) { // 遍历key集合
                String key = (String) iterator.next(); // 得到key
//                label = new Label(column++, 0, key); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值
                label = new Label(0,++rownum, key); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值

                sheet.addCell(label); // 将单元格加到页
            }
            // 遍历jsonArray
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject item = jsonArray.getJSONObject(i); // 得到数组的每项
                iterator = item.keys(); // 得到key集合
//                column = 0;// 从第0列开始放
                rownum=0;//从第0行开始放
                while (iterator.hasNext()) {
                    String key = iterator.next(); // 得到key
                    String value = item.getString(key); // 得到key对应的value
//                    label = new Label(column++,(i + 1), value); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值
                    label = new Label((i+1),++rownum, value); // 第一个参数是单元格所在列,第二个参数是单元格所在行,第三个参数是值

                    sheet.addCell(label); // 将单元格加到页
                }
            }
            writableWorkbook.write(); // 加入到文件中
            writableWorkbook.close(); // 关闭文件，释放资源
        } catch (Exception e) {
//            result.put("result", "failed"); // 将调用该函数的结果返回
//            result.put("reason", e.getMessage()); // 将调用该函数失败的原因返回
//            return result;
            e.printStackTrace();
        }
        result.put("result", "successed");
        return result;
    }

    public static void main(String[] args) {
        String testJsonString="{ \"data\": [ { \"error_code\": 0, \"reason\": \"ok\", \"result\": { \"staffNumRange\": \"-\", \"updatetime\": 1581787483117, \"fromTime\": 1511712000000, \"type\": 1, \"holderlist\": [ { \"amount\": 100, \"toco\": 2, \"id\": 3076688783, \"logo\": \"https://img5.tianyancha.com/logo/lll/44eecec75d180727e02359d4a9ec02a2.png@!f_200x200\", \"alias\": \"顺风尚德\", \"name\": \"厦门顺风尚德光电系统集成有限公司\", \"capitalActl\": [ ], \"type\": 1, \"capital\": [ { \"amomon\": \"100万人民币\", \"time\": \"2067-11-23\", \"percent\": \"100.00%\", \"paymet\": \"\" } ] } ], \"isMicroEnt\": 1, \"id\": 3121287883, \"percentileScore\": 5839, \"regNumber\": \"350203200831998\", \"regCapital\": \"100万人民币\", \"name\": \"厦门顺风德聚新能源科技有限公司\", \"regInstitute\": \"厦门市思明区市场监督管理局\", \"regLocation\": \"厦门市思明区南投路3号1306A\", \"industry\": \"科技推广和应用服务业\", \"approvedTime\": 1565884800000, \"socialStaffNum\": 0, \"logo\": \"https://img5.tianyancha.com/logo/lll/a22f33c959587792a802c9f7263780de.png@!f_200x200\", \"taxNumber\": \"91350203MA2YQWCT4E\", \"businessScope\": \"合同能源管理；其他未列明科技推广和应用服务业；燃气、太阳能及类似能源家用器具制造；集成电路设计；太阳能光伏系统施工；经营各类商品和技术的进出口（不另附进出口商品目录），但国家限定公司经营或禁止进出口的商品及技术除外；电气设备批发；其他未列明批发业（不含需经许可审批的经营项目）；其他机械设备及电子产品批发；管道和设备安装；太阳能发电；节能技术推广服务；其他未列明专业技术服务业（不含需经许可审批的事项）；其他技术推广服务；电力供应；工程和技术研究和试验发展；信息系统集成服务；信息技术咨询服务。\", \"alias\": \"厦门顺风\", \"orgNumber\": \"MA2YQWCT4\", \"estiblishTime\": 1511712000000, \"regStatus\": \"存续\", \"legalPersonName\": \"芮永胜\", \"toTime\": 3089462400000, \"legalPersonId\": 2341010085, \"sourceFlag\": \"http://qyxy.baic.gov.cn/\", \"actualCapital\": \"-\", \"flag\": 1, \"correctCompanyId\": \"\", \"companyOrgType\": \"有限责任公司（自然人投资或控股的法人独资）\", \"updateTimes\": 1581787482000, \"base\": \"fj\", \"companyType\": 0, \"creditCode\": \"91350203MA2YQWCT4E\", \"companyId\": 141871430 } } ] }";
        //        JSONArray testJsonArray=JSONArray.fromObject(testJsonString);
        String testJsonString2="{\"result\":[ { \"staffNumRange\": \"-\", \"updatetime\": 1581787483117, \"fromTime\": 1511712000000, \"type\": 1, \"holderlist\": [ { \"amount\": 100, \"toco\": 2, \"id\": 3076688783, \"logo\": \"https://img5.tianyancha.com/logo/lll/44eecec75d180727e02359d4a9ec02a2.png@!f_200x200\", \"alias\": \"顺风尚德\", \"name\": \"厦门顺风尚德光电系统集成有限公司\", \"capitalActl\": [ ], \"type\": 1, \"capital\": [ { \"amomon\": \"100万人民币\", \"time\": \"2067-11-23\", \"percent\": \"100.00%\", \"paymet\": \"\" } ] } ], \"isMicroEnt\": 1, \"id\": 3121287883, \"percentileScore\": 5839, \"regNumber\": \"350203200831998\", \"regCapital\": \"100万人民币\", \"name\": \"厦门顺风德聚新能源科技有限公司\", \"regInstitute\": \"厦门市思明区市场监督管理局\", \"regLocation\": \"厦门市思明区南投路3号1306A\", \"industry\": \"科技推广和应用服务业\", \"approvedTime\": 1565884800000, \"socialStaffNum\": 0, \"logo\": \"https://img5.tianyancha.com/logo/lll/a22f33c959587792a802c9f7263780de.png@!f_200x200\", \"taxNumber\": \"91350203MA2YQWCT4E\", \"businessScope\": \"合同能源管理；其他未列明科技推广和应用服务业；燃气、太阳能及类似能源家用器具制造；集成电路设计；太阳能光伏系统施工；经营各类商品和技术的进出口（不另附进出口商品目录），但国家限定公司经营或禁止进出口的商品及技术除外；电气设备批发；其他未列明批发业（不含需经许可审批的经营项目）；其他机械设备及电子产品批发；管道和设备安装；太阳能发电；节能技术推广服务；其他未列明专业技术服务业（不含需经许可审批的事项）；其他技术推广服务；电力供应；工程和技术研究和试验发展；信息系统集成服务；信息技术咨询服务。\", \"alias\": \"厦门顺风\", \"orgNumber\": \"MA2YQWCT4\", \"estiblishTime\": 1511712000000, \"regStatus\": \"存续\", \"legalPersonName\": \"芮永胜\", \"toTime\": 3089462400000, \"legalPersonId\": 2341010085, \"sourceFlag\": \"http://qyxy.baic.gov.cn/\", \"actualCapital\": \"-\", \"flag\": 1, \"correctCompanyId\": \"\", \"companyOrgType\": \"有限责任公司（自然人投资或控股的法人独资）\", \"updateTimes\": 1581787482000, \"base\": \"fj\", \"companyType\": 0, \"creditCode\": \"91350203MA2YQWCT4E\", \"companyId\": 141871430 }]}";
        JSONObject testJsonObject=JSONObject.fromObject(testJsonString);
        ExcelUtil.createExcel("C:\\Users\\85114\\Desktop\\TestExcel\\test.xls",testJsonObject,"data");
        JSONObject testJsonObject2=JSONObject.fromObject(testJsonString2);
        ExcelUtil.createExcel("C:\\Users\\85114\\Desktop\\TestExcel\\test2.xls",testJsonObject2,"result");
        String testJsonString3="{\"holderlist\": [ { \"amount\": 100, \"toco\": 2, \"id\": 3076688783, \"logo\": \"https://img5.tianyancha.com/logo/lll/44eecec75d180727e02359d4a9ec02a2.png@!f_200x200\", \"alias\": \"顺风尚德\", \"name\": \"厦门顺风尚德光电系统集成有限公司\", \"capitalActl\": [ ], \"type\": 1, \"capital\": [ { \"amomon\": \"100万人民币\", \"time\": \"2067-11-23\", \"percent\": \"100.00%\", \"paymet\": \"\" } ] } ]}";
        JSONObject testJsonObject3=JSONObject.fromObject(testJsonString3);
        ExcelUtil.createExcel("C:\\Users\\85114\\Desktop\\TestExcel\\test3.xls",testJsonObject3,"holderlist");
        String testJsonString4="{\"capital\": [ { \"amomon\": \"100万人民币\", \"time\": \"2067-11-23\", \"percent\": \"100.00%\", \"paymet\": \"\" } ]}";
        JSONObject testJsonObject4=JSONObject.fromObject(testJsonString4);
        ExcelUtil.createExcel("C:\\Users\\85114\\Desktop\\TestExcel\\test4.xls",testJsonObject4,"capital");

    }
}
