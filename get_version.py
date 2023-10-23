import requests
import re
import sys
import warnings

warnings.filterwarnings("ignore")

# Reference:https://docs.microsoft.com/en-us/Exchange/new-features/build-numbers-and-release-dates

def get_exchange_version(build_number):
    strlist = build_number.split('.')
    if len(strlist) < 3:
        return "Unknown Exchange Server"

    if int(strlist[0]) == 4:
        return 'Exchange Server 4.0'

    elif int(strlist[0]) == 5:
        if int(strlist[1]) == 0:
            return 'Exchange Server 5.0'
        elif int(strlist[1]) == 5:
            return 'Exchange Server 5.5'

    elif int(strlist[0]) == 6:
        if int(strlist[1]) == 0:
            return 'Exchange Server 2000'
        elif int(strlist[1]) == 5:
            if int(strlist[2]) == 6944:
                return 'Exchange Server 2003'
            elif int(strlist[2]) == 7226:
                return 'Exchange Server 2003 SP1'
            elif int(strlist[2]) == 7683:
                return 'Exchange Server 2003 SP2'
            elif int(strlist[2]) == 7653:
                return 'Exchange Server 2003 post-SP2'
            elif int(strlist[2]) == 7654:
                return 'Exchange Server 2003 post-SP2'
            elif int(strlist[1]) == 0:
                return 'Exchange 2000 Server'

    elif int(strlist[0]) == 8:
        if int(strlist[1]) == 0:
            return 'Exchange Server 2007'
        elif int(strlist[1]) == 1:
            return 'Exchange Server 2007 SP1'
        elif int(strlist[1]) == 2:
            return 'Exchange Server 2007 SP2'
        elif int(strlist[1]) == 3:
            return 'Exchange Server 2007 SP3'

    elif int(strlist[0]) == 14:
        if int(strlist[1]) == 0:
            return 'Exchange Server 2010'
        elif int(strlist[1]) == 1:
            return 'Exchange Server 2010 SP1'
        elif int(strlist[1]) == 2:
            return 'Exchange Server 2010 SP2'
        elif int(strlist[1]) == 3:
            return 'Exchange Server 2010 SP3'

    elif int(strlist[0]) == 15:
        if int(strlist[1]) == 0:
            if int(strlist[2]) == 516:
                return "Exchange Server 2013 RTM"
            elif int(strlist[2]) == 620:
                return "Exchange Server 2013 CU1"
            elif int(strlist[2]) == 712:
                return "Exchange Server 2013 CU2"
            elif int(strlist[2]) == 775:
                return "Exchange Server 2013 CU3"
            elif int(strlist[2]) == 847:
                return "Exchange Server 2013 SP1"
            elif int(strlist[2]) == 913:
                return "Exchange Server 2013 CU5"
            elif int(strlist[2]) == 995:
                return "Exchange Server 2013 CU6"
            elif int(strlist[2]) == 1044:
                return "Exchange Server 2013 CU7"
            elif int(strlist[2]) == 1076:
                return "Exchange Server 2013 CU8"
            elif int(strlist[2]) == 1104:
                return "Exchange Server 2013 CU9"
            elif int(strlist[2]) == 1130:
                return "Exchange Server 2013 CU10"
            elif int(strlist[2]) == 1156:
                return "Exchange Server 2013 CU11"
            elif int(strlist[2]) == 1178:
                return "Exchange Server 2013 CU12"
            elif int(strlist[2]) == 1210:
                return "Exchange Server 2013 CU13"
            elif int(strlist[2]) == 1236:
                return "Exchange Server 2013 CU14"
            elif int(strlist[2]) == 1263:
                return "Exchange Server 2013 CU15"
            elif int(strlist[2]) == 1293:
                return "Exchange Server 2013 CU16"
            elif int(strlist[2]) == 1320:
                return "Exchange Server 2013 CU17"
            elif int(strlist[2]) == 1347:
                return "Exchange Server 2013 CU18"
            elif int(strlist[2]) == 1365:
                return "Exchange Server 2013 CU19"
            elif int(strlist[2]) == 1367:
                return "Exchange Server 2013 CU20"
            elif int(strlist[2]) == 1395:
                return "Exchange Server 2013 CU21"
            elif int(strlist[2]) == 1473:
                return "Exchange Server 2013 CU22"
            elif int(strlist[2]) == 1497:
                return "Exchange Server 2013 CU23"

        elif int(strlist[1]) == 1:
            if int(strlist[2]) == 225:
                return "Exchange Server 2016 RTM"
            elif int(strlist[2]) == 396:
                return "Exchange Server 2016 CU1"
            elif int(strlist[2]) == 466:
                return "Exchange Server 2016 CU2"
            elif int(strlist[2]) == 544:
                return "Exchange Server 2016 CU3"
            elif int(strlist[2]) == 669:
                return "Exchange Server 2016 CU4"
            elif int(strlist[2]) == 845:
                return "Exchange Server 2016 CU5"
            elif int(strlist[2]) == 1034:
                return "Exchange Server 2016 CU6"
            elif int(strlist[2]) == 1261:
                return "Exchange Server 2016 CU7"
            elif int(strlist[2]) == 1415:
                return "Exchange Server 2016 CU8"
            elif int(strlist[2]) == 1466:
                return "Exchange Server 2016 CU9"
            elif int(strlist[2]) == 1531:
                return "Exchange Server 2016 CU10"
            elif int(strlist[2]) == 1591:
                return "Exchange Server 2016 CU11"
            elif int(strlist[2]) == 1713:
                return "Exchange Server 2016 CU12"
            elif int(strlist[2]) == 1779:
                return "Exchange Server 2016 CU13"
            elif int(strlist[2]) == 1847:
                return "Exchange Server 2016 CU14"
            elif int(strlist[2]) == 1913:
                return "Exchange Server 2016 CU15"
            elif int(strlist[2]) == 1979:
                return "Exchange Server 2016 CU16"
            elif int(strlist[2]) == 2044:
                return "Exchange Server 2016 CU17"
            elif int(strlist[2]) == 2106:
                return "Exchange Server 2016 CU18"
            elif int(strlist[2]) == 2176:
                return "Exchange Server 2016 CU19"
            elif int(strlist[2]) == 2242:
                return "Exchange Server 2016 CU20"
            elif int(strlist[2]) == 2308:
                return "Exchange Server 2016 CU21"
            elif int(strlist[2]) == 2375:
                return "Exchange Server 2016 CU22"
            elif int(strlist[2]) == 2507:
                return "Exchange Server 2016 CU23"

        elif int(strlist[1]) == 2:
            if int(strlist[2]) == 196:
                return "Exchange Server 2019 Preview"
            elif int(strlist[2]) == 221:
                return "Exchange Server 2019 RTM"
            elif int(strlist[2]) == 330:
                return "Exchange Server 2019 CU1"
            elif int(strlist[2]) == 397:
                return "Exchange Server 2019 CU2"
            elif int(strlist[2]) == 464:
                return "Exchange Server 2019 CU3"
            elif int(strlist[2]) == 529:
                return "Exchange Server 2019 CU4"
            elif int(strlist[2]) == 595:
                return "Exchange Server 2019 CU5"
            elif int(strlist[2]) == 659:
                return "Exchange Server 2019 CU6"
            elif int(strlist[2]) == 721:
                return "Exchange Server 2019 CU7"
            elif int(strlist[2]) == 792:
                return "Exchange Server 2019 CU8"
            elif int(strlist[2]) == 858:
                return "Exchange Server 2019 CU9"
            elif int(strlist[2]) == 922:
                return "Exchange Server 2019 CU10"
            elif int(strlist[2]) == 986:
                return "Exchange Server 2019 CU11"
            elif int(strlist[2]) == 1118:
                return "Exchange Server 2019 CU12"
            elif int(strlist[2]) == 1258:
                return "Exchange Server 2019 CU13"


def get_exchange_build_number(url):
    try:
        r = requests.get(url, verify=False)
        match = re.search(r'/owa/auth/(\d+\.\d+\.\d+(\.\d+)*)/', r.text)
        if match:
            return match.group(1)

    except Exception as e:
        print('[X] Error: %s' % e)

    return ""


def get_exchange_info(url):
    build_number = get_exchange_build_number(url)
    if len(build_number):
        version = get_exchange_version(build_number)
        print(version)
        print('Build number: %s' % build_number)
    else:
        print("[!] Build number not found.")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('[X] Wrong parameter\n')
        print('Usage:')
        print('\t%s <url>' % (sys.argv[0]))
        print('\nExample:')
        print('\t%s https://exchange.test.com/owa' % (sys.argv[0]))
        sys.exit(0)
    else:
        get_exchange_info(sys.argv[1])
