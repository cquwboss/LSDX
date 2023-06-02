import math
import os

import numpy as np
import xlwings as xw
from scipy.stats import pearsonr
from sklearn.metrics import mean_squared_error, mean_absolute_error
from symfit import parameters, variables, sin, cos, Fit
import numpy as np
import matplotlib.pyplot as plt


def read_trend_data(data_path, sheet_name):
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
    # 读取第一列数据
    all_data = sht.range('b2').expand('table')
    rows = all_data.rows.count
    all_data = np.array(sht.range(f'b2:b{rows+1}').value).reshape(1, -1)
    # print(all_data)
    # 每次读写完毕，记得关闭
    wb.close()
    app.quit()
    return all_data


def fourier_series(x, f, n=0):
    """
    Returns a symbolic fourier series of order `n`.

    :param n: Order of the fourier series.
    :param x: Independent variable
    :param f: Frequency of the fourier series
    """
    # Make the parameter objects for all the terms
    a0, *cos_a = parameters(','.join(['a{}'.format(i) for i in range(0, n + 1)]))
    sin_b = parameters(','.join(['b{}'.format(i) for i in range(1, n + 1)]))
    # Construct the series
    series = a0 + sum(ai * cos(i * f * x) + bi * sin(i * f * x)
                      for i, (ai, bi) in enumerate(zip(cos_a, sin_b), start=1))
    return series

def ploy_series(x,n=0):
    a0, *cos_a = parameters(','.join(['a{}'.format(i) for i in range(0, n + 1)]))
    series = a0 + sum(ai*np.power(x,i) for i, ai  in enumerate(cos_a, start=1))
    return series

def get_model_ploy_series(curve_count):
    x, y = variables('x, y')
    model_dict = {y: ploy_series(x,  n=curve_count)}
    print(model_dict)
    return model_dict

def get_model_fourier_series(curve_count):
    x, y = variables('x, y')
    w, = parameters('w')
    model_dict = {y: fourier_series(x, f=w, n=curve_count)}
    print(model_dict)
    return model_dict


def write_params_to_excel(data_path, sheet_name, trend_data, params,train_data,test_data,predict_data):
    """
    将分解重构后的数据写入excel文件
    :param data_path: 新文件路径
    :param sheet_name: 工作表名
    :param data: 原始数据
    :param X_ssa: 经过ssa变化后的数据
    :param recon_data: 重构数据
    :param mse: 计算的均方误差
    :param mae: 绝对误差
    :param pr: 相关系数
    :return:
    """
    """
       将趋势数据\系数单独写入文件
       :param file_path: 文件路径
       :param sht_name: 表名
       :param trend_data: (61,)
       :param results: 系数array 其大小根据逼近次数来定
       """
    if os.path.exists(data_path):
        print("已有相同文件，将会覆盖此文件")
        os.remove(data_path)  # 删除文件
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()  # 创建新工作簿
    sht = wb.sheets.add(sheet_name)  # 创建工作表

    titles = np.array(['ssa趋势数据', '拟合系数','训练数据','测试数据','预测数据'])  # 所有标题
    sht.range('A1').value = titles  # 写入标题
    sht.range('A2').value = trend_data.T  # 写入趋势数据
    sht.range('B2').value = params.reshape(1,-1).T  # 写入参数公式
    sht.range('C2').value = train_data.reshape(1,-1).T  # 写入训练数据
    sht.range('D2').value = test_data.reshape(1,-1).T  # 写入测试数据
    sht.range('E2').value = predict_data.reshape(1,-1).T  # 写入预测数据

    sht.range('A1:E1').column_width = 25  # 调整列宽
    sht.range('B1').column_width = 75  # 调整列宽
    wb.save(data_path)  # 保存excel 文件
    wb.close()
    app.quit()


def cal_error_index(data, recon_data):
    mse = mean_squared_error(data, recon_data)
    mae = mean_absolute_error(data, recon_data)
    pr = pearsonr(data, recon_data)
    print("均方误差：", mse, "平均绝对误差：", mae, "相关系数：", pr)
    return mse, mae, pr



if __name__ == '__main__':
    # 傅里叶逼近次数 或者多项式项数
    curve_count = 3
    # prece 前百分之多少作为拟合数据集，后面的作为测试
    prece = 0.8 #
    trend_data = read_trend_data('ssa_result.xls', 'ssa_result')
    train_len = math.ceil(trend_data.shape[1] * prece)  # 用来拟合的训练集长度

    x = np.arange(0, trend_data.shape[1], 1)
    #  间隔设定1，相当于1s
    x_train_data = np.arange(0, train_len, 1)
    # 趋势训练数据，转为1维
    y_train_data = trend_data[0][:train_len]
    # 测试集的x
    x_test_data = np.arange(train_len, trend_data.shape[1], 1)
    # 趋势测试数据，转为1维
    y_test_data = trend_data[0][train_len:]
    # 使用numpy自带的拟合函数。多项式拟合好过 傅里叶级数拟合
    z = np.polyfit( x_train_data, y_train_data, curve_count)  # 用多项式拟合

    # 获取拟合后的多项式 ，
    p = np.poly1d(z)
    print(p)  # 在屏幕上打印拟合多项式
    # 计算拟合后的y值
    y_test_predict = p(x_test_data)
    # 计算误差
    mse, mae, pr = cal_error_index(y_test_data, y_test_predict)
    # 存放系数
    para_result = []
    para_result.append(str(p))
    # np拟合系数组合
    for i,value in enumerate(p.coefficients):
        para_result.append("a{}:{}".format(i,value))  # 组合一下,类似 a0:2
    # 写入excel
    write_params_to_excel('trend_curve_data.xls', 'trend_curve', trend_data, np.array(para_result),y_train_data,y_test_data,y_test_predict)
    #
    # 画出结果对比
    plt.plot(x_train_data, y_train_data)
    # plt.plot(x_train_data, fit.model(x=x_train_data, **fit_result.params).y, color='green', ls=':')
    plt.plot(x_train_data, p(x_train_data), color='green', ls=':')
    plt.plot(x_test_data, y_test_data,color='blue',) # test 部分
    plt.plot(x_test_data, y_test_predict, color='red', ls=':')
    plt.show()

    # 寻优
    # mses = []
    # maes = []
    # trend_data = read_trend_data('ssa_result.xls', 'ssa_result')
    # for curve_count in range(6,7):
    #     for prece in np.linspace(0.90,0.90,1):
    #         # 读取趋势数据
    #         print('***************cc:{}***************pre:{}'.format(curve_count,prece))
    #
    #         train_len = math.ceil(trend_data.shape[1]*prece) # 用来拟合的训练集长度
    #         x = np.arange(0, trend_data.shape[1], 1)
    #         #  间隔设定1，相当于1s
    #         x_train_data = np.arange(0, train_len, 1)
    #         # 趋势训练数据，转为1维
    #         y_train_data = trend_data[0][:train_len]
    #         # 测试集的x
    #         x_test_data = np.arange(train_len, trend_data.shape[1], 1)
    #         # 趋势测试数据，转为1维
    #         y_test_data = trend_data[0][train_len:]
    #
    #         z = np.polyfit( x_train_data, y_train_data, curve_count)  # 用3次多项式拟合
    #
    #         # 获取拟合后的多项式
    #         p = np.poly1d(z)
    #         print(p)  # 在屏幕上打印拟合多项式
    #
    #         # 计算拟合后的y值
    #         y_test_predict = p(x_test_data)
    #
    #         # 获取模型
    #         # 傅里叶级数拟合
    #         # model_dict = get_model_fourier_series(curve_count)
    #         # 多项式拟合
    #         # model_dict = get_model_ploy_series(curve_count)
    #         # # 适配模型
    #         # fit = Fit(model_dict, x=x_train_data, y=y_train_data)
    #         # # fit = Fit(model_dict, x=x, y=trend_data[0])
    #         # # 开始执行，循环 找到 parameters
    #         # fit_result = fit.execute()
    #         # # 使用系数和公式预测
    #         # y_test_predict = fit.model(x=x_test_data, **fit_result.params).y
    #         # 计算误差
    #         mse, mae, pr = cal_error_index(y_test_data,y_test_predict)
    #         mses.append(mse)
    #         maes.append(mae)
    #         # 打印系数
    #         # print(fit_result.params.items())
    #
    # min_mse = np.min(np.array(mses))
    # min_mae = np.min(np.array(maes))
    # print("最小的mse:",min_mse,"最小的mae",min_mae)

