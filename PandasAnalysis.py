import pandas as pd

def to_xlsx(portsdf: object) -> None:
    packetsdf = packets(portsdf)
    print(packetsdf)
    speedsdf = speeds(portsdf)
    print(speedsdf)
    lldp_avgdf = LLDP_packets(portsdf)
    print(lldp_avgdf)
    with pd.ExcelWriter('Pandas_analysis.xlsx') as writer:
        packetsdf.to_excel(writer, sheet_name='Сумма пакетов', index=False)
        speedsdf.to_excel(writer, sheet_name='Скорости передачи пакетов', index=False)
        lldp_avgdf.to_excel(writer, sheet_name='Среднее количество пакетов', index=False)
        portsdf.to_excel(writer, sheet_name='Общая информация', index=False)

def packets(portsdf: object) -> object:
    sumin = portsdf['Packets in'].sum()
    sumout = portsdf['Packets out'].sum()
    packetsdict = {'Sum in': [sumin], 
                   'Sum out': [sumout]}
    return pd.DataFrame(packetsdict)

def speeds(portsdf: object) -> object:
    max_in = []
    max_out = []
    min_in = []
    min_out = []
    for i in range(4):
        if (i % 2 == 0):
            str = 'On'
        else:
            str = 'Off'
        if (i < 2):
            status = 'Oper '
        else:
            status = 'Admin '
        buffdf = portsdf[portsdf[status + 'status'] == str]
        max_in.append(buffdf['Actual speed in'].max())
        max_out.append(buffdf['Actual speed out'].max())
        min_in.append(buffdf['Actual speed in'].min())
        min_out.append(buffdf['Actual speed out'].min())
    speedsdict = {'Oper status': ['On', 'Off','',''], 
                  'Admin status': ['','','On','Off'], 
                  'Max speed in': max_in, 
                  'Max speed out': max_out, 
                  'Min speed in': min_in, 
                  'Min speed out': min_out}
    return pd.DataFrame(speedsdict)

def LLDP_packets(portsdf: object) -> object:
    buffdf = portsdf[portsdf['LLDP UP'] == 'On']
    avg_up_in = round(buffdf['Packets in'].mean())
    avg_up_out = round(buffdf['Packets out'].mean())
    buffdf = portsdf[portsdf['LLDP UP'] == 'Off']
    avg_down_in = round(buffdf['Packets in'].mean())
    avg_down_out = round(buffdf['Packets out'].mean())
    packetsdict = {'LLDP UP': ['On', 'On', 'Off', 'Off'], 
                   'Packets': ['In', 'Out', 'In', 'Out'], 
                   'Average': [avg_up_in, avg_up_out, avg_down_in, avg_down_out]}
    return pd.DataFrame(packetsdict)

def main():
    portsdf = pd.read_csv('all_info.csv')
    print(portsdf)

    to_xlsx(portsdf)
    
if __name__ == "__main__":
    main()