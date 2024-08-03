from datetime import timedelta, datetime

class TimeInterval:
    def __init__(self):
        self.start: datetime = None
        self.end: datetime = None
        self.duration: timedelta = None

    def __str__(self) -> str:
        if self.start is None or self.end is None:
            return "None"
        return f"{self.start.strftime('%I:%M%p')} - {self.end.strftime('%I:%M%p')}"

    def set_interval(self, start: str, end: str | None = None):
        if end:
            self.set_start(start)
            self.set_end(end)
        else:
            start, end = start.split(" - ")
            self.set_start(start)
            self.set_end(end)
        self.duration: timedelta = self.end - self.start

    def set_interval_dt(self, start: datetime, end: datetime):
        self.start = start
        self.end = end
        self.duration: timedelta = self.end - self.start

    def set_start(self, start: datetime):
        self.start = start

    def set_start(self, start: str):
        self.start = datetime.strptime(start, "%I:%M%p")

    def set_end(self, end: datetime):
        self.end = end

    def set_end(self, end: str):
        self.end = datetime.strptime(end, "%I:%M%p")


class Day:
    def __init__(self, name, day_type="REGULAR"):
        self._name: str = None
        self.name: str = name
        self._day_type: str = None
        self.day_type: str = day_type
        self.interval: TimeInterval = TimeInterval()

    @property
    def day_type(self):
        return self._day_type

    @day_type.setter
    def day_type(self, value):
        if value not in ["DIA DE DESCANSO", "VACACIONES", "PAGO HORAS FERIADO", "REGULAR", None]:
            raise ValueError("Invalid day type")
        self._day_type = value

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        if value not in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo", None]:
            raise ValueError("Invalid day name")
        self._name = value

    def __str__(self) -> str:
        return f"{self.name} - {self.day_type} - {self.interval}"
    
    def to_dict(self):
        return {
            "day_type": self.day_type,
            "start": self.interval.start.strftime("%I:%M%p") if self.interval.start else None,
            "end": self.interval.end.strftime("%I:%M%p") if self.interval.end else None,
        }

    def set_interval(self, interval: str):
        if self.day_type == "REGULAR":
            self.interval.set_interval(interval)
        else:
            raise ValueError("Interval can only be set for regular days")

    def get_duration(self):
        return self.interval.duration

    def set_name(self, name: str):
        self.name = name


class Schedule:
    def __init__(self):
        self.monday : Day = Day("Lunes")
        self.tuesday : Day = Day("Martes")
        self.wednesday : Day = Day("Miércoles")
        self.thursday : Day = Day("Jueves")
        self.friday : Day = Day("Viernes")
        self.saturday : Day = Day("Sábado")
        self.sunday : Day = Day("Domingo")

    def get_day_by_index(self, index):
        if index == 0:
            return self.monday
        if index == 1:
            return self.tuesday
        if index == 2:
            return self.wednesday
        if index == 3:
            return self.thursday
        if index == 4:
            return self.friday
        if index == 5:
            return self.saturday
        if index == 6:
            return self.sunday

    def __str__(self) -> str:
        return f"{self.monday}\n{self.tuesday}\n{self.wednesday}\n{self.thursday}\n{self.friday}\n{self.saturday}\n{self.sunday}"


class Employee:
    def __init__(self):
        self.name = None
        self._category: str = None
        self.category: str = None
        self.schedule = Schedule()

    @property
    def category(self):
        return self._category

    @category.setter
    def category(self, value):
        if value not in ["RS", "SELF", "CAJERO", None]:
            raise ValueError(f"Invalid category {value}")
        self._category = value

    def __str__(self) -> str:
        return f"{self.name} - {self.category}"
    
    def to_dict(self):
        return {
            "name": self.name,
            "monday": self.schedule.monday.to_dict(),
            "tuesday": self.schedule.tuesday.to_dict(),
            "wednesday": self.schedule.wednesday.to_dict(),
            "thursday": self.schedule.thursday.to_dict(),
            "friday": self.schedule.friday.to_dict(),
            "saturday": self.schedule.saturday.to_dict(),
            "sunday": self.schedule.sunday.to_dict()
        }