# This is a sample Python script.
import os
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pyts.decomposition import SingularSpectrumAnalysis
from scipy.fftpack import fft, fftfreq, ifft

from sklearn.metrics import mean_squared_error  # 均方误差
from sklearn.metrics import mean_absolute_error  # 平方绝对误差
from scipy.stats import pearsonr
from sklearn.metrics import r2_score  # R square


def read_shift_data(data_path, sheet_name):
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
    # 读取时间序列数据,并reshape成二维，方便pyts运算
    all_data = np.array(sht.range('b1').expand('table').value).reshape(1, -1)
    # print(all_data)
    # 每次读写完毕，记得关闭
    wb.close()
    app.quit()
    return all_data


def pyts_ssa(data, window_size_L, groups_t):
    """
    使用pyts库ssa算法解构 data
    :param data: 原始数据
    :param window_size_L: 窗口长度
    :param groups_t: 分组个数
    :return: 返回分解后的信号
    """
    # Singular Spectrum Analysis
    ssa = SingularSpectrumAnalysis(window_size=window_size_L, groups=groups_t)
    X_ssa = ssa.fit_transform(data)
    return X_ssa


def reconstruct_signal(X_ssa, reconstruct_count):
    """
    重构信息
    :param X_ssa: ssa后的信号
    :param reconstruct_count: 重构信号需要包含 ssa分量的 个数
    """

    assert X_ssa.shape[0] >= reconstruct_count  # 断言必须分解后的信号要大于你设置选择ssa分量的个数
    recon_s = np.zeros((1, X_ssa.shape[1]))  # 初始化 recon_s 其维度 (1,64)
    for i in range(reconstruct_count):
        recon_s += X_ssa[i].reshape(1, -1)  # 根据所需分量个数进行累加
    return recon_s




def write_ssa_data(data_path, sheet_name, data, X_ssa, recon_data, indexes):
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
    if os.path.exists(data_path):
        print("已有相同文件，将会覆盖此文件")
        os.remove(data_path)  # 删除文件
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()  # 创建新工作簿
    sht = wb.sheets.add(sheet_name)  # 创建工作表
    all_data = np.vstack((data[0], X_ssa[0], X_ssa[1], recon_data[0])).T  # 将数据组合在一起
    titles = np.array(['原始数据', 'ssa趋势数据', 'ssa周期数据', 'ssa重构数据', '指标参数'])  # 所有标题
    sht.range('A1').value = titles  # 写入标题
    sht.range('A2').value = all_data  # 写入所有数据
    sht.range('E2').value = indexes  # 写入指标参数
    sht.range('A1:D1').column_width = 15
    sht.range('E1').column_width = 30  # 调整该单元格长度，显示全部信息
    wb.save(data_path)  # 保存excel 文件
    wb.close()
    app.quit()


def show_result(data, X_ssa, recon_data, restruct_sig_fft, w_L, g_t, r_c):
    """
    显示所有数据信息图片，并保存
    :param data: 原始数据
    :param X_ssa: 经过ssa变换后的数据，包含了5个分量
    :param recon_data: 重构后的信号，叠加前4个分量
    :param restruct_sig_fft: 经过 ifft 重构后的信号
    :param w_L: 窗口长度
    :param g_t: 分组数
    :param r_c: 重构信号用到前多少分量
    """
    # 画原始信号
    plt.figure(figsize=(16, 12))
    ax1 = plt.subplot(221)
    ax1.plot(data[0], 'o-', label='Original')
    ax1.legend(loc='best', fontsize=14)

    # 画ssa分解后的分量信号
    ax2 = plt.subplot(222)
    for i in range(X_ssa.shape[0]):
        ax2.plot(X_ssa[i], 'o--', label='SSA {0}'.format(i + 1))
    ax2.legend(loc='best', fontsize=14)

    # 趋势信号
    ax3 = plt.subplot(223)
    ax3.plot(X_ssa[0], 'o-', label='Trend Signal')
    ax3.legend(loc='best', fontsize=14)

    # 周期信号
    ax4 = plt.subplot(224)
    ax4.plot(X_ssa[1], 'o-', label='Circle Signal')
    ax4.legend(loc='best', fontsize=14)
    plt.suptitle('Singular Spectrum Analysis', fontsize=20)
    plt.tight_layout()
    plt.subplots_adjust(top=0.88)
    plt.savefig(str(w_L) + '_' + str(g_t) + '_' + str(r_c) + '.png', dpi=600)

    # 画ssa重构信号与原始信号的对比
    plt.figure(figsize=(16, 12))
    ax1 = plt.subplot(121)
    ax1.plot(data[0], 'o-', label='Original')
    ax1.legend(loc='best', fontsize=14)

    # 画ssa分解后的分量信号
    ax2 = plt.subplot(122)
    ax2.plot(recon_data[0], 'o--', label='Reconstruction Signal(4 ssa)')
    ax2.legend(loc='best', fontsize=14)
    plt.suptitle('Reconstruction Signal(ssa)', fontsize=20)
    plt.tight_layout()
    plt.subplots_adjust(top=0.88)
    plt.savefig(str(w_L) + '_' + str(g_t) + '_' + str(r_c) + 'Reconstruction Signal.png', dpi=600)

    # 画fft重构信号与趋势信号的对比
    # 画趋势信号
    plt.figure(figsize=(16, 12))
    ax1 = plt.subplot(121)
    ax1.plot(X_ssa[0], 'o-', label='Trend')
    ax1.legend(loc='best', fontsize=14)
    # 画ifft 重构信号
    ax2 = plt.subplot(122)
    ax2.plot(restruct_sig_fft, 'o--', label='Reconstruction Signal(ifft)')
    ax2.legend(loc='best', fontsize=14)
    plt.suptitle('Reconstruction Signal(ifft)', fontsize=20)
    plt.tight_layout()
    plt.subplots_adjust(top=0.88)
    plt.savefig(str(w_L) + '_' + str(g_t) + '_' + str(r_c) + 'Reconstruction Signal(ifft).png', dpi=600)
    # plt.show()


def fit_wave_by_fft(data):
    """
    使用 fft 拟合 趋势项
    :param data: 趋势项数据
    """
    sig_fft = fft(data, len(data))
    amp = np.abs(sig_fft)  # 获取转换后的信号幅度
    max_amp = np.max(amp)  # 利用ifft重构信号,寻找最大、最小幅值，根据此来设定阈值，舍弃一些无用信号
    min_amp = np.min(amp)
    amp_yu = min_amp + 0.0008 * (max_amp - min_amp)  # 阈值设定 小于阈值幅度对应的频率舍弃
    hig_freq = sig_fft.copy()  # 复制转换信号
    hig_freq[amp < amp_yu] = 0  # 将小于阈值幅度的信号舍弃
    res_signal = np.abs(ifft(hig_freq))  # 重构信号，输出实部
    # plt.figure()
    # plt.plot(data,'o--')
    # plt.plot(res_signal, 'o-')
    # plt.grid()
    # plt.show()
    return res_signal


def cal_error_index(data, recon_data):
    mse = mean_squared_error(data, recon_data)
    mae = mean_absolute_error(data, recon_data)
    pr = pearsonr(data[0], recon_data[0])
    print("均方误差：", mse, "平均绝对误差：", mae, "相关系数：", pr)
    return mse, mae, pr



if __name__ == '__main__':
    data = read_shift_data('YY208.xls', 'YY208')  # 读取excel信号
    # 最佳参数  window_L = 28 , groups_t = 5, reconstruct_count = 4
    window_L = 9  # 窗口长度
    groups_t = 5  # 分组数量，也代表了分解后的信号个数
    reconstruct_count = 5 # 取分解后的分量的个数来重组信号
    X_ssa = pyts_ssa(data, window_L, groups_t)  # ssa分解
    recon_data = reconstruct_signal(X_ssa, reconstruct_count)  # 重构信号，使用了前4种分量
    restruct_sig_fft = fit_wave_by_fft(X_ssa[0])  # 傅里叶拟合趋势项，返回经过ifft重构后的信号
    mse, mae, pr = cal_error_index(data, recon_data)  # 计算误差指标

    show_result(data, X_ssa, recon_data,restruct_sig_fft, window_L, groups_t, reconstruct_count)  # 显示结果,以图片保存
    # 组合指标
    indexes = '窗口长度：' + str(window_L) + '\n分组数：' + str(groups_t) + '\n重构ssa分量数：' + str(
        reconstruct_count) + '\n均方误差:' + str(mse) + '\n绝对误差:' + str(mae) + '\n相关系数:' + str(pr[0])
    write_ssa_data('ssa_result.xls', 'ssa_result', data, X_ssa, recon_data, indexes)

    #  这段注释代码通过比较误差寻找最优的参数组合，找到最优组合为 window_L = 9 , groups_t = 5, reconstruct_count = 5
    # mses = []
    # maes = []
    # for window_L in range(5,30):
    #     for groups_t in range(3,6):
    #         for reconstruct_count in range(2,groups_t):
    #             print("L:",window_L,"t:",groups_t,"c:",reconstruct_count)
    #             X_ssa = pyts_ssa(data, window_L, groups_t)  # ssa分解
    #             recon_data = reconstruct_signal(X_ssa, reconstruct_count)  # 重组信号
    #             #fit_wave_by_fft(X_ssa[0])
    #             mse,mae,pr = cal_error_index(data,recon_data)
    #             mses.append(mse)
    #             maes.append(mae)
    #             show_result(data, X_ssa, recon_data,window_L,groups_t,reconstruct_count)  # 显示结果
    #
    # min_mse = np.min(np.array(mses))
    # min_mae = np.min(np.array(maes))
    # print("最小的mse:",min_mse,"最小的mae",min_mae)
