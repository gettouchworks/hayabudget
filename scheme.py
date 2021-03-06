#coding=utf-8 
fields = [
u"月份",
u"公司名称",
u"工资发放地",
u"实际工作地",
u"费用类别",
u"费用分类",
u"预算单位",
u"计薪方式",
u"各体系负责人",
u"考核单元负责人",
u"状态",
u"一级部门",
u"二级部门",
u"人数",
u"员工编号",
u"姓名",
u"性别",
u"岗位名称",
u"职级",
u"在职月数",
u"基本工资",
u"绩效工资",
u"岗位工资",
u"固定加班工资",
u"其他工资",
u"工龄工资",
u"房屋津贴",
u"误工费",
u"市场误餐费",
u"午餐补贴",
u"其他津贴",
u"项目制奖金",
u"成本节约考核奖",
u"6S奖",
u"创新奖",
u"特殊类人员奖金",
u"其他奖金",
u"绩效奖金",
u"奖金池",
u"平日加班工资",
u"周末加班工资",
u"法定假日加班工资",
u"专项奖金",
u"经济补偿金",
u"社保基数",
u"养老保险",
u"失业保险",
u"医疗保险",
u"工伤保险",
u"生育保险",
u"特殊保险",
u"住房公积金数据",
u"住房公积金",
u"正常午餐补助",
u"加班餐补",
u"定额交通报销",
u"定额汽油",
u"定额通讯补助",
u"探亲路费",
u"妇女节礼品费",
u"中秋节日费",
u"体检费",
u"医疗费",
u"员工活动经费",
u"生日费",
u"年会费",
u"年度评优费用",
u"庆典",
u"理想家平台费",
u"羽毛球",
u"运动会",
u"比赛竞赛",
u"其他活动",
u"存档费",
u"结婚喜金",
u"生育礼金",
u"管理人员房租",
u"其他补助",
u"职称考试评定",
u"企业培训",
u"职工教育经费其他",
u"工会经费",
u"劳保用品",
u"工服",
u"防暑降温",
u"劳动保护其他"
]

query_base = [
'工资',
'奖金',
'补偿',
'社保',
'住房',
'福利',
'教育',
'工会',
'劳保'
]
query_extend = [
'餐补',
'交通',
'通信',
'探亲',
'节日',
'医疗',
'固定活动',
'其他活动',
'其他福利',
'教育',
'工会',
]
querys = {}
querys["工资"] = fields[20:42]
querys["奖金"] = [fields[42]]
querys["补偿"] = [fields[43]]
querys["社保"] = fields[45:51]
querys["住房"] = [fields[52]]
querys['餐补'] = fields[53:55]
querys['交通'] = fields[55:57]
querys['通信'] = [fields[57]]
querys['探亲'] = [fields[58]]
querys['节日'] = fields[59:61]
querys['医疗'] = fields[61:63]
querys['固定活动'] = fields[63:67]
querys['其他活动'] = fields[67:73]
querys['其他福利'] = fields[73:78]
querys['教育'] = fields[78:81]
querys['工会'] = [fields[81]]
querys['劳保'] = fields[82:86]
querys["福利"] = fields[53:78]

addons = [
u"月份",
u"公司名称",
u"工资发放地",
u"实际工作地",
u"费用类别",
u"费用分类",
u"预算单位",
u"计薪方式",
u"各体系负责人",
u"考核单元负责人",
u"状态",
u"一级部门",
u"二级部门",
u"人数",
u"员工编号",
u"姓名",
u"性别",
u"岗位名称",
u"职级",
u"在职月数",
u"月社保基数增加额",
u"社保基数合规月数",
u"月养老增加额",
u"月工伤增加额",
u"月失业增加额",
u"月生育增加额",
u"月医疗增加额",
u"社保合规总额",
u"住房基数增加额",
u"合规月数",
u"月住房公积金增加额",
u"住房合规总额",
u"社保稽查补缴"
]

monthList = [u"1月",u"2月",u"3月",u"4月",u"5月",u"6月",u"7月",u"8月",u"9月",u"10月",u"11月",u"12月"]

if __name__ == "__main__":
	for i in fields[20:41]:
		print i

# monthList = [u"一月",u"二月",u"三月",u"四月",u"五月",u"六月",u"七月",u"八月",u"九月",u"十月",u"十一月",u"十二月"]