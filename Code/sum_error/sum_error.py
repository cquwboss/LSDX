import math
import os

import numpy as np
import xlwings as xw
from scipy.stats import pearsonr
from sklearn.metrics import mean_squared_error, mean_absolute_error
from symfit import parameters, variables, sin, cos, Fit
import numpy as np
import matplotlib.pyplot as plt


def read_write_result_data(data_path, sheet_name):
    """
    读取excel文件中的数据信息，以array形式返回
    :param data_path: 源数据文件路径
    :param sheet_name: 数据表名
    :return: 读取到的数据
    """
    # 判断是否有此文件
    if not os.path.exists(data_path):
        print("无此文件")
        return
    # 新建工作簿 (如果不接下一条代码的话，Excel只会一闪而过，卖个萌就走了）
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(data_path)
    # 打开工作表
    sht = wb.sheets[sheet_name]
    # 读取时间序列数据
    all_data = np.array(sht.range('a3').expand('table').value)[:,:5]
    assert all_data.shape[1] == 5
    # 读取趋势项数据,
    trend_data,pso_circle,gsa_circle,pso_gsa_circle,recon = all_data[:,0],all_data[:,1],all_data[:,2],all_data[:,3],all_data[:,4]

    # 计算pso 累积误差
    params = [] #所有指标
    pso_mse, pso_mae, pso_pr = cal_error_index(recon, trend_data+pso_circle)
    pso_params = {
        'mse': pso_mse,
        'mae': pso_mae,
        'pr': pso_pr
    }
    params.append(str(pso_params))
    # 计算gsa
    gsa_mse, gsa_mae, gsa_pr = cal_error_index(recon, trend_data + gsa_circle)
    gsa_params = {
        'mse': gsa_mse,
        'mae': gsa_mae,
        'pr': gsa_pr
    }
    params.append(str(gsa_params))
    # 计算pso-gsa指标
    pso_gsa_mse, pso_gsa_mae, pso_gsa_pr = cal_error_index(recon, trend_data + pso_gsa_circle)
    pso_gsa_params = {
        'mse': pso_gsa_mse,
        'mae': pso_gsa_mae,
        'pr': pso_gsa_pr
    }
    params.append(str(pso_gsa_params))
    # 每次读写完毕，记得关闭
    # 写入标题
    sht.range('f2').expand('table').value = ['pso评估指标','gsa评估指标','pso-gsa评估指标']
    sht.range('f3').expand('table').value = params
    sht.range('f1:h1').column_width = 75  # 调整列宽
    wb.save()
    wb.close()
    app.quit()
    return all_data


def cal_error_index(data, recon_data):
    mse = mean_squared_error(data, recon_data)
    mae = mean_absolute_error(data, recon_data)
    pr = pearsonr(data, recon_data)
    print("均方误差：", mse, "平均绝对误差：", mae, "相关系数：", pr)
    return mse, mae, pr



if __name__ == '__main__':
    # 读写累加误差
    read_write_result_data('累加误差评价.xlsx', 'Sheet1')
