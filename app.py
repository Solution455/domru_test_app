import datetime
import tkinter as tk

import paramiko
import speedtest
import wmi

#
class Speedtest:
    # TODO: откалибровать замеры скорости, протестировать апи на разных пк
    """
    Замер скорости
    """
    _st = speedtest.Speedtest()

    def speed_info(self):
        self._st.get_best_server()
        download_speed = self._st.download()
        upload_speed = self._st.upload()
        ping = self._st.results.ping
        return {'Download': download_speed / 1_000_000, 'Upload': upload_speed / 1_000_000, 'Ping': ping}


class FtpConnect:
    # TODO: настроить отправку и подключение к актуальному фтп компании
    """
    Подключение к фтп, передача информации на сервер.
    """

    def connect(self):
        transport = paramiko.Transport(('#', 22))
        transport.connect(username='#', password='#')
        sftp = paramiko.SFTPClient.from_transport(transport)

        sftp.put('temp_domrutracker.txt',
                 'tracker/' + f"tracker-{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
        sftp.close()
        transport.close()


class WMI:
    """
    Сбор информации с пк
    """
    _c = wmi.WMI()

    def _get_cpu_info(self):
        cpu = self._c.Win32_Processor()[0]
        return {'cpu_name': cpu.Name, 'cpu_cores': cpu.NumberOfCores, 'cpu_virtcores': cpu.NumberOfLogicalProcessors}

    def _get_mem_info(self):
        mem = self._c.Win32_PhysicalMemory()
        mem_capacity = sum((int(el.Capacity) for el in mem)) // (1024 * 1024)
        mem_speed = sum((int(el.Speed) for el in mem)) // len(mem)
        return {'mem_capacity': mem_capacity, 'mem_speed': mem_speed}

    def _get_disk_info(self):
        disk = self._c.Win32_LogicalDisk()
        dct = {}
        for key, el in enumerate(disk, start=1):
            free_space_in_mb = int(el.freespace) // (1024 * 1024)
            if free_space_in_mb < 100:
                dct[f'disk_size{key}'] = el.size
                dct[f'disk_free_space{key}'] = el.freespace
        return dct

    def _get_net_info(self):
        net = self._c.Win32_NetworkAdapterConfiguration(IPEnabled=True)
        netstat = {}
        for index, el in enumerate(net, start=1):
            netstat[f'name{index}'] = el.Description
            netstat[f'ip{index}'] = el.IPAddress
            netstat[f'mac{index}'] = el.MACAddress
            netstat[f'netmask{index}'] = el.IPSubnet
            netstat[f'gateway{index}'] = el.DefaultIPGateway
            netstat[f'dns{index}'] = el.DNSServerSearchOrder
        return netstat

    def _get_os_info(self):
        _os = self._c.Win32_OperatingSystem()[0]
        return {'os': _os.Caption}

    def _get_motherboard_info(self):
        mother = self._c.Win32_BaseBoard()[0]
        return {'mother_manufacturer': mother.Manufacturer, 'mother_product': mother.Product}

    def _get_video_info(self):
        gpu = self._c.Win32_VideoController()[0]
        return {'gpu_name': gpu.Name, 'gpu_ver': gpu.DriverVersion}

    def _get_info(self):
        if self._get_disk_info():
            return {'cpu_info': self._get_cpu_info(), 'mem_info': self._get_mem_info(),
                    'disk_info': self._get_disk_info(),
                    'motherbd': self._get_motherboard_info(),
                    'video_info': self._get_video_info(), 'net_info': self._get_net_info(),
                    'osinfo': self._get_os_info()}
        else:
            return {'cpu_info': self._get_cpu_info(), 'mem_info': self._get_mem_info(),
                    'motherbd': self._get_motherboard_info(),
                    'video_info': self._get_video_info(), 'net_info': self._get_net_info(),
                    'osinfo': self._get_os_info()}


class Interpreter(WMI):
    # TODO: Визуальный интерфейс
    """
    Структурирование и запись информации в файл для передачи на FTP
    """
    __st = Speedtest()
    __wmi = WMI()
    filename = f"tracker-{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"

    def _take_info(self):
        return {'pc_info': self.__wmi._get_info(), 'st': self.__st.speed_info()}

    def converter(self):
        start_time = datetime.datetime.now()
        result = ''
        info = self._take_info()
        st_info = info['st']
        trouble_counter = 0

        cpu_info = info['pc_info']['cpu_info']
        cpu_cores = cpu_info['cpu_cores']
        cpu_virtcores = cpu_info['cpu_virtcores']
        cpu_name = cpu_info['cpu_name']
        if int(st_info['Download']) <= 40 and int(st_info['Upload']) <= 40:
            result += f'•[BAD] Возможно, проблемы с интернет соединением: ' \
                      f'\n    └ Download: {st_info["Download"]:.2f} Мбит/с' \
                      f'\n     └ Upload: {st_info["Upload"]:.2f} Мбит/с' \
                      f'\n       └ Ping: {st_info["Ping"]:.2f} мс\n'
            trouble_counter += 1
        else:
            result += f'•[OK] Соединение в норме:' \
                      f'\n    └ Download: {st_info["Download"]:.2f} Мбит/с' \
                      f'\n     └ Upload: {st_info["Upload"]:.2f} Мбит/с' \
                      f'\n       └ Ping: {st_info["Ping"]:.2f} мс\n'

        if cpu_cores < 2 and cpu_virtcores < 4:
            result += '•[BAD] Процессор слабый.\n'
            trouble_counter += 1
        else:
            result += f'•[OK] Процессор в норме: {cpu_name} \n    └ Кол-во виртуальных ядер: ' \
                      f'{cpu_virtcores}\n     └ Кол-во физ. ядер: {cpu_cores}\n'

        mem_info = info['pc_info']['mem_info']
        mem_capacity = mem_info['mem_capacity']
        mem_speed = mem_info['mem_speed']
        if mem_capacity < 2000 and int(mem_speed) < 1066:
            result += f'•[BAD] Свободной памяти мало:\n' \
                      f'      └ Всего RAM: {mem_capacity:.2f} МБ\n' \
                      f'        └ Частота: {int(mem_speed)} МГц\n'
            trouble_counter += 1
        else:
            result += f'•[OK] Память в норме:\n' \
                      f'    └ Всего RAM: {mem_capacity:.2f} МБ\n' \
                      f'      └ Частота: {int(mem_speed)} МГц\n'

        net_info = info['pc_info']['net_info']
        wired_net = [el for el in net_info.values() if el is not None and 'Ethernet' in el]
        wired_net = [str(el) for el in wired_net]
        if wired_net:
            result += f'•[OK] Подключение по кабелю через:\n' \
                      f'     └ {" или ".join(wired_net)}\n'
        else:
            result += '•[BAD] Нет кабельного подключения\n'
            trouble_counter += 1

        if info['pc_info'].get('disk_info'):
            disk_info = info['pc_info']['disk_info']
            disk_info = [f"{int(el) / (1024 ** 2):.2f} MB" for el in disk_info.values() if el is not None]
            if disk_info:
                result += f'[BAD]• Недостаточно памяти на одном из дисков: {", ".join(disk_info)}\n'
                trouble_counter += 1
        else:
            result += '•[OK] На дисках достаточно места.\n'

        motherbd = info['pc_info']['motherbd']
        motherbd = [el for el in motherbd.values() if el is not None]
        motherbd = [str(el) for el in motherbd]
        if motherbd:
            result += f'•[OK] Материнская плата: {", ".join(motherbd)}\n'
        else:
            result += '•[BAD] Нет материнской платы\n'
            trouble_counter += 1

        video_info = info['pc_info']['video_info']
        video_info = [el for el in video_info.values() if el is not None]
        video_info = [str(el) for el in video_info]
        if video_info:
            result += f'•[OK] Видеокарта: {video_info[0]}\n' \
                      f'    └ Версия драйвера: {video_info[1]}\n{"=" * 100}\n\n'
        else:
            result += f'•[BAD] Нет видеокарты\n{"=" * 100}\n\n'
            trouble_counter += 1

        end_time = datetime.datetime.now()
        execution_time = end_time - start_time
        total_seconds = int(execution_time.total_seconds())
        minutes, seconds = divmod(total_seconds, 60)

        if trouble_counter == 0:
            result += '•[OK] Все проверки пройдены.\n'
        elif trouble_counter <= 2:
            result += '•[ATTENTION] Возможны сложности.\n'
        else:
            result += '•[BAD] Пк слабый.\n'

        result += f'• Время выполнения: {minutes}:{seconds}\n'
        with open('temp_domrutracker.txt', "w", encoding='utf-8') as file:
            file.write(result)

        return result


def main(obj: Interpreter, ftp: FtpConnect):
    """Начальная точка старта приложения"""
    print('Программа запущена')
    result = obj.converter()
    ftp.connect()
    print('Данные отправлены')
    return result


def visual_view():
    def start_button_click():
        result = main(Interpreter(), FtpConnect())
        console_text.insert(tk.END, result)
        console_text.insert(tk.END, "Выполнение программы завершено\n")

    root = tk.Tk()
    root.title("Дом.ру трекер")
    root.geometry("680x450")
    root.resizable(False, False)
    root.iconbitmap("qd.ico")

    console_text = tk.Text(root, width=100, height=25)
    console_text.insert('end', 'Нажмите "Старт" для проверки компьютера на соответствие удаленной работы.\n'
                               'Программа на время выполнения подвиснет, это нормально.\n\n\n\n\n')
    console_text.pack()

    forward_button = tk.Button(root, text="Старт", command=start_button_click)
    forward_button.pack()

    root.mainloop()


if __name__ == '__main__':
    visual_view()
