#!/usr/bin/env python
# coding=utf-8


from sonarqube import SonarQubeClient
import excel_w


url = 'http://192.168.50.202:9000'
username = "admin"
password = "1234.com"
token = "12e0a4a6602d76bab7fb17c8e5aab12a1c2b7e64"
#myprod_token = "599bccc489347292fcc72952e7b573c4606e75bd"


class mysonar(object):
    def __init__(self, sonar):
        """需要传入sonar连接的实例化对象"""
        self.sonar = sonar

    def get_project_key(self):
        """查询所有的项目"""
        project_key_list = list(self.sonar.projects.search_projects())
        project_list = []
        for pl in project_key_list:
            project_list.append(pl['key'])
        return project_list

    def get_prod_filter(self, project_list):
        """按照prod关键字过滤出生产环境中的代码扫描项目"""
        project_prod_filter = filter(lambda x:'prod' in x, project_list)
        return project_prod_filter
        

    def get_project_data(self, projectkey):
        """获取项目的扫描数据，存入列表中"""
        component = self.sonar.measures.get_component_with_specified_measures(
            component=projectkey, metricKeys="bugs,vulnerabilities,code_smells,ncloc", branch=None)
        measuers_list = component['component']['measures']
        temp_map = {}
        project_data_list = []
        for metrics in measuers_list:
            metric_type = metrics['metric']
            metric_value = metrics['value']
            temp_map[metric_type] = metric_value
        # print(temp_map)
        project_data_list.insert(0, projectkey)
        project_data_list.insert(1, temp_map['bugs'])
        project_data_list.insert(2, temp_map['vulnerabilities'])
        project_data_list.insert(3, temp_map['code_smells'])
        project_data_list.insert(4, temp_map['ncloc'])
        return project_data_list

    def get_issues_data(self, projectkey, issue_type_list):
        """将每个issues类型的结果存入列表中"""
        severities_list = ['BLOCKER', 'CRITICAL', 'MAJOR', 'MINOR', 'INFO']
        res_data_list = []
        for i in issue_type_list:
            for s in severities_list:
                issue_list = list(self.sonar.issues.search_issues(componentKeys=projectkey, branch='master', severities=s, types=i,resolved='false'))
                #str_res = projectkey + ":" + i + ":" + s + ":"+str(len(issue_list))
                str_res = str(len(issue_list))
                res_data_list.append(str_res)

        return res_data_list

    def clear_issues_data(self, res_data_list):
        """按照列表中的类型排序，进行数据清洗，便于写入excel中"""
        bug_list = res_data_list[0:5]
        vulnerability_list = res_data_list[5:10]
        code_smell_list = res_data_list[10:15]
        all_issue_data_list = []
        all_issue_data_list.append("-".join(bug_list))
        all_issue_data_list.append("-".join(vulnerability_list))
        all_issue_data_list.append("-".join(code_smell_list))
        return all_issue_data_list

                


def write_data_to_excel():
    """将数据写入Excel中"""
    sonar = SonarQubeClient(sonarqube_url=url, token=token)
    #severities_list = ['BLOCKER', 'CRITICAL', 'MAJOR', 'MINOR', 'INFO']
    issue_type_list = ['BUG','VULNERABILITY','CODE_SMELL']
    my = mysonar(sonar)
    myprojcet_list = my.get_project_key()
    my_prod_filter = my.get_prod_filter(myprojcet_list)
    #print(my_prod_filter)
    res_project_data = []
    for p in my_prod_filter:
        res = my.get_project_data(p)
        res_data_list = my.get_issues_data(p, issue_type_list)
        clear_resuelt = my.clear_issues_data(res_data_list)
        resuelt = res + clear_resuelt
        res_project_data.append(resuelt)
    #print(res_project_data)
    myexcle = excel_w.write_table_row(res_project_data)
    myexcle.wirte_data()



if __name__ == '__main__':
    write_data_to_excel()


    

 




