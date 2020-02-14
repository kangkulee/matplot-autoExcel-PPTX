import Test01.create_csv as cc
import Test01.make_avgSnr as avgSnr
import Test01.make_rssi as rssi
import Test01.make_snr as snr

if __name__ == '__main__':
    cc.create_csv()
    avgSnr.make_avgSnr()
    rssi.make_rssi()
    snr.make_snr()
