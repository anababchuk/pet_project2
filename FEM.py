import pandas as pd
import numpy as np
import json
from pandas import DataFrame, ExcelWriter
df = DataFrame()
workbook = pd.ExcelFile('optimizator.xlsx')
d = {}
for sheet_name in workbook.sheet_names:
    df = workbook.parse(sheet_name)
    d[sheet_name] = df
df.columns = ['Скважина', 'Куст', 'Частота УЭЦН было, Гц', 'Частота УЭЦН стало, Гц',
              'Расход газлифтного газа было, тыс.м3/сут', 'Расход газлифтного газа стало, тыс.м3/сут',
              'Потребляемая мощность было, кВт∙ч', 'Потребляемая мощность стало, кВт∙ч',
              'Линейное давление было, атм', 'Линейное давление стало, атм',
              'Забойное давление было, атм', 'Забойное давление стало, атм',
              'Дебит жидкости было, м3/сут', 'Дебит жидкости стало, м3/сут',
              'Дебит нефти было, т/сут', 'Дебит нефти стало, т/сут',
              'Дебит воды было, м3/сут', 'Дебит воды стало, м3/сут',
              'Дебит газа было, тыс.м3/сут', 'Дебит газа стало, тыс.м3/сут',
              'NPV было, тыс.руб/сут', 'NPV стало, тыс.руб/сут',
              'LC было, тыс.руб/сут', 'LC стало, тыс.руб/сут', 'Обводненность, %']
df.info()

#уберем вторую строку с названиями колонок
idx=df.index
dfx=df
dfx=dfx.drop(dfx.index[0])
dfx.index=range(len(dfx))
dfx.tail()

#подготовим данные для функции расчета экономических показателей
a=dfx[dfx.columns[12]]
liquid_arr=a.values

a=dfx[dfx.columns[24]]
wtc_arr=a.values

a=dfx[dfx.columns[6]]
power_arr=a.values

a=dfx[dfx.columns[18]]
gas_arr=a.values

with open('fem_data.json') as f:
    fem_data = json.load(f)

def calculate_econ(fem_data, liquid_arr, wtc_arr, power_arr, gas_arr):
    """
        Расчет экономических показателей по вкаждой скважине
        Args:
            fem_data: словарь с экономическими данными
            liquid_arr: массив с дебитами жидкости поскважинно
            wtc_arr: массив с обводненностью поскважинно
            power_arr: массив с мощностями поскважинно
            gas_arr: массив с дебитами газа поскважинно

        Returns NPV, NPV/тонна, lifting costs
    """
    oil_arr = liquid_arr * (1 - wtc_arr / 100)
    water_arr = liquid_arr * wtc_arr / 100

    ndpi = np.asarray([fem_data['NDPI'] * i for i in oil_arr])
    income_oil = np.asarray([fem_data['netbackOil'] * i *
                             (1 - fem_data['oilLost'] / 100) for i in oil_arr])
    income_gas = np.asarray([fem_data['netbackGas'] * i *
                             fem_data['gasRealization'] / 100000 for i in gas_arr])

    expense_ee_1 = np.asarray([fem_data['specificEnergyConsumption_PPD'] * i for i in liquid_arr])
    expense_ee_2 = np.asarray([fem_data['specificEnergyConsumption_InputLiquidTransport'] * i for i in liquid_arr])
    expense_ee_3 = np.asarray([(fem_data['specificEnergyConsumption_OilTransport'] + fem_data[
        'specificEnergyConsumption_OilPreparation'] + fem_data[
                                    'specificEnergyConsumption_OutputOilTransport']) * i * (1 - fem_data['oilLost'])
                               for i in oil_arr])
    expense_ee_4 = np.asarray([fem_data['specificEnergyConsumption_WaterTransport'] * i for i in water_arr])
    expense_ee_5 = np.asarray([j * 24 for i, j in zip(liquid_arr, power_arr)])

    expense_chemistry_1 = np.asarray([i * fem_data['otherExpenseOil'] * (1 - fem_data['oilLost']) for i in oil_arr])
    expense_chemistry_2 = np.asarray([i * fem_data['otherExpensePPD'] for i in liquid_arr])

    other_expense_gas_1 = np.asarray([i / 1000 * fem_data['gasVariableExpense'] for i in gas_arr])
    other_expense_gas_2 = np.asarray(
        [i / 1000 * fem_data['gasBurning'] * fem_data['gasBurningFine'] for i in gas_arr])

    income_sum = income_oil + income_gas
    expense_ee_oil_sum = (expense_ee_1 + expense_ee_2 + expense_ee_3 + expense_ee_4 + expense_ee_5) * fem_data[
        'energyCost']
    expense_chemistry_sum = expense_chemistry_1 + expense_chemistry_2
    expense_ee_gas_sum = np.asarray(
        [fem_data['energyCost'] * i * fem_data['specificEnergyConsumption_Gas'] / 1000 for i in gas_arr])
    other_expense_gas_sum = other_expense_gas_1 + other_expense_gas_2

    lc = expense_ee_oil_sum + expense_ee_gas_sum + other_expense_gas_sum + expense_chemistry_sum + (
            fem_data['espOperationCost'] + fem_data['espHireCost'] + fem_data['tubingOperationCost'] +
            fem_data['extraEquipmentOperationCost'])

    marginal_revenue = income_sum - ndpi - lc
    marginal_revenue = marginal_revenue.astype(np.float32)
    oil_arr = oil_arr.astype(np.float32)
    marginal_revenue_per_tonne = np.divide(marginal_revenue, oil_arr, out=np.zeros_like(marginal_revenue),
                                           where=oil_arr != 0)
    if marginal_revenue.all() > 0:
        marginal_revenue = marginal_revenue * (1 - fem_data['incomeTax'] / 100)
        marginal_revenue = marginal_revenue.astype(np.float32)
        oil_arr = oil_arr.astype(np.float32)
        marginal_revenue_per_tonne = np.divide(marginal_revenue, oil_arr, out=np.zeros_like(marginal_revenue),
                                               where=oil_arr != 0)
    else:
        pass

    return [marginal_revenue, marginal_revenue_per_tonne, lc]

#записываем результаты расчетов в датафрейм, добавив номера скважин
result=calculate_econ(fem_data, liquid_arr, wtc_arr, power_arr, gas_arr)
resultFrame = pd.DataFrame(data=result, index=['marginal_revenue', 'marginal_revenue_per_tonne', 'lc'])
resultFrame = resultFrame.transpose()
resultFrame.index=dfx[dfx.columns[0]]

#записываем результаты в Эксель и сохраняем
writer=ExcelWriter('МП,ЛК.xlsx')
resultFrame.to_excel(writer, 'Результаты')
writer.save()