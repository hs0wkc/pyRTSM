from psychrolib import SetUnitSystem, SI, CalcPsychrometricsFromTWetBulb, CalcPsychrometricsFromRelHum
from pyMEP import Quantity
from pyMEP.hvac.internal_heat_gains import *
from pyMEP.hvac.external_heat_gains import *
from pyMEP.hvac.coolingload import *
from pyMEP.hvac.lighting import *
from pyMEP.hvac.people import *
from pyMEP.hvac.equipment import *
from lib._resource import *

Q_ = Quantity

class ComfortZone:

	_oreintation : Quantity
	_ventilation : Quantity

	def __init__(self, id: str, weather: WeatherData) -> None:
		self.ID = id
		self.weather_data = weather
		self.Width : Quantity = Q_(0, 'm')
		self.Length : Quantity = Q_(0, 'm')
		self.Height : Quantity = Q_(0, 'm')
		self.SpaceType : int
		self.safety : Quantity
		self.max_hr : list

		self.Roof : Roof = Roof(id='Roof')
		self.Ceiling : Ceiling = Ceiling(id='Ceiling')
		self.Floor : Floor = Floor(id='Floor')
		self.Wall_A : Wall = Wall(id='Wall_A')
		self.Wall_B : Wall = Wall(id='Wall_B')
		self.Wall_C : Wall = Wall(id='Wall_C')
		self.Wall_D : Wall = Wall(id='Wall_D')

		self.Light_HeatGain = LightingHeatGain(ID="Lighting_0")
		self.Light_HeatGain.add_lighting(SpaceLighting.create(ID="ls0", schedule=self.Light_HeatGain.usage_schedule))
		self.People_HeatGain = PeopleHeatGain('People_0')
		self.Equipment_HeatGain = EquipmentHeatGain('Equipment_0')
		self.Equipment_HeatGain.add_equipment(GenericAppliance.create(ID='eqp0', schedule=self.Equipment_HeatGain.usage_schedule))

		self.external_load_df = pd.DataFrame([i for i in range(24)], columns=['Hr'])
		self.external_load_df.set_index('Hr', inplace=True)
		self.internal_load_df = self.external_load_df.copy()
		self.ventilation_load_df = self.external_load_df.copy()
		self.wall_load_df = self.external_load_df.copy()
		self.window_load_df = self.external_load_df.copy()
		self.cooling_load_df = self.external_load_df.copy()

	def Calculate(self) -> None:
		# UPDATE WEATHER DATA
		self.weather_data.update_sun_position()
		self.weather_data.synthetic_daily_db_profiles()
		match self.SpaceType:
			case 0:		# Single Floor
				self.Roof.IsEnabled = True
				self.Floor.IsEnabled = False
			case 1:		# Highest Floor
				self.Roof.IsEnabled = True
				self.Floor.IsEnabled = True
			case 2:		# Middle Floor
				self.Roof.IsEnabled = False
				self.Floor.IsEnabled = True
			case 3:		# Lowest Floor
				self.Roof.IsEnabled = False
				self.Floor.IsEnabled = False
		self.Ceiling.IsEnabled = not self.Roof.IsEnabled
		Setting.NRTS_zones = 'Exterior' if self.Wall_A.wall_type.value or self.Wall_B.wall_type.value or self.Wall_C.wall_type.value or self.Wall_D.wall_type.value else 'Interior'
		self.ns_rts_zone = RTS.rts_values(nrts=True, zones = Setting.NRTS_zones ,
													 room_construction = Setting.NRTS_Room_construction,
													 carpet = Setting.NRTS_Carpet,
													 glass = Setting.NRTS_Glass)
		self.s_rts_zone = RTS.rts_values(nrts=False, room_construction = Setting.NRTS_Room_construction,
													 carpet = Setting.NRTS_Carpet,
													 glass = Setting.NRTS_Glass)
		# ROOF
		if not self.Roof.IsEnabled:
			self.Roof.cooling_load_df = None
			self.external_load_df['Roof'] = 0.0
		else:
			self.Roof.Update_ns_rts(self.ns_rts_zone)
			self.external_load_df['Roof'] = self.Roof.cooling_load_df['TOTAL_CL']
		# FLOOR
		self.Floor.Update_ns_rts(self.ns_rts_zone)
		self.Ceiling.cooling_load_df = self.Floor.cooling_load_df.copy()
		if not self.Floor.IsEnabled:
			self.Floor.cooling_load_df = None
			self.external_load_df['Floor'] = 0.0
		else:
			self.external_load_df['Floor'] = self.Floor.cooling_load_df['TOTAL_CL']
		# CEILING
		if not self.Ceiling.IsEnabled:
			self.Ceiling.cooling_load_df = None
			self.external_load_df['Ceiling'] = 0.0
		else:
			self.external_load_df['Ceiling'] = self.Ceiling.cooling_load_df['TOTAL_CL']
		# WALL
		self.Wall_A.Update_ns_rts(self.ns_rts_zone)
		self.wall_load_df['Wall-A'] = self.Wall_A.cooling_load_df['TOTAL_CL']
		self.Wall_B.Update_ns_rts(self.ns_rts_zone)
		self.wall_load_df['Wall-B'] = self.Wall_B.cooling_load_df['TOTAL_CL']
		self.Wall_C.Update_ns_rts(self.ns_rts_zone)
		self.wall_load_df['Wall-C'] = self.Wall_C.cooling_load_df['TOTAL_CL']
		self.Wall_D.Update_ns_rts(self.ns_rts_zone)
		self.wall_load_df['Wall-D'] = self.Wall_D.cooling_load_df['TOTAL_CL']
		# WINDOW
		if not self.Wall_A.windows:
			self.window_load_df['Win-A'] = 0.0
		else:
			window = self.Wall_A.windows.get('Win-A')
			window.S_RTS = self.s_rts_zone
			window.update_cooling_load()
			self.window_load_df['Win-A'] = window.cooling_load_df['TOTAL_CL']
		if not self.Wall_B.windows:
			self.window_load_df['Win-B'] = 0.0
		else:
			window = self.Wall_B.windows.get('Win-B')
			window.S_RTS = self.s_rts_zone
			window.update_cooling_load()
			self.window_load_df['Win-B'] = window.cooling_load_df['TOTAL_CL']
		if not self.Wall_C.windows:
			self.window_load_df['Win-C'] = 0.0
		else:
			window = self.Wall_C.windows.get('Win-C')
			window.S_RTS = self.s_rts_zone
			window.update_cooling_load()
			self.window_load_df['Win-C'] = window.cooling_load_df['TOTAL_CL']
		if not self.Wall_D.windows:
			self.window_load_df['Win-D'] = 0.0
		else:
			window = self.Wall_D.windows.get('Win-D')
			window.S_RTS = self.s_rts_zone
			window.update_cooling_load()
			self.window_load_df['Win-D'] = window.cooling_load_df['TOTAL_CL']
		# LIGHTING
		self.Light_HeatGain.Update_ns_rts(self.ns_rts_zone)
		self.internal_load_df['Lighting'] = self.Light_HeatGain.cooling_load_df['TOTAL_CL']
		# PEOPLE
		self.People_HeatGain.Update_ns_rts(self.ns_rts_zone)
		self.internal_load_df['People'] = self.People_HeatGain.cooling_load_df['TOTAL_CL']
		# EQUIPMENT
		self.Equipment_HeatGain.Update_ns_rts(self.ns_rts_zone)
		self.internal_load_df['Equipment'] = self.Equipment_HeatGain.cooling_load_df['TOTAL_CL']
		# VENTILATION
		qs = np.round(self._cs * self._ventilation * (np.array([i.m for i in self.weather_data.T_db_prof]) - Setting.Inside_DB.m), 0)
		self.ventilation_load_df['SHG'] =  qs * np.array(self.Light_HeatGain.usage_profile)
		SetUnitSystem(SI)
		hum_ratio_o = CalcPsychrometricsFromTWetBulb(self.weather_data.T_db_des.m, self.weather_data.T_wb_mc.m, 101325)[0] * 1000
		hum_ratio_i = CalcPsychrometricsFromRelHum(Setting.Inside_DB.m, Setting.Inside_RH, 101325)[0] * 1000
		ql = np.round(self._cl * self._ventilation * (hum_ratio_o - hum_ratio_i), 0)
		self.ventilation_load_df['LHG'] =  ql * np.array(self.Light_HeatGain.usage_profile)
		self.ventilation_load_df['TOTAL_CL'] = self.ventilation_load_df['SHG'] + self.ventilation_load_df['LHG']
		# SUMMARY
		self.wall_load_df['TOTAL_CL'] = 0.0
		self.wall_load_df['TOTAL_CL'] = self.wall_load_df.sum(axis=1)
		self.external_load_df['Wall'] = self.wall_load_df['TOTAL_CL']
		self.window_load_df['TOTAL_CL'] = 0.0
		self.window_load_df['TOTAL_CL'] = self.window_load_df.sum(axis=1)
		self.external_load_df['Window'] = self.window_load_df['TOTAL_CL']
		self.external_load_df['TOTAL_CL'] = 0.0
		self.external_load_df['TOTAL_CL'] = self.external_load_df.sum(axis=1)
		self.internal_load_df['TOTAL_CL'] = 0.0
		self.internal_load_df['TOTAL_CL'] = self.internal_load_df.sum(axis=1)
		self.cooling_load_df['external'] = self.external_load_df['TOTAL_CL']
		self.cooling_load_df['internal'] = self.internal_load_df['TOTAL_CL']
		self.cooling_load_df['ventilation'] = self.ventilation_load_df['TOTAL_CL']
		self.cooling_load_df['TOTAL_CL'] = 0.0
		self.cooling_load_df['TOTAL_CL'] = self.cooling_load_df.sum(axis=1)
		total_cl = self.cooling_load_df['TOTAL_CL']
		self.max_hr = total_cl[total_cl == total_cl.max()].index.tolist()
		print('---------- CALCULATED ------------')

	# region
	@property
	def Oreintation(self) -> Quantity:
		return self._oreintation
	@Oreintation.setter
	def Oreintation(self, v: int) -> None:
		# ASHRAE Fundamentals 2021, Chapter 14, §14.11 surface azimuth
		v = -180 if v==180 else v;
		self._oreintation = Q_(v, 'deg')
		A = 180 - v if v>=0 else -(180+v)
		self.Wall_A.psi = Q_(A, 'deg')
		B = 90 - v if v>=-90 else -(270+v)
		self.Wall_B.psi = Q_(B, 'deg')
		C = -v
		self.Wall_C.psi = Q_(C, 'deg')
		D = -(90 + v) if -180<=v<90 else 270 - v
		self.Wall_D.psi = Q_(D, 'deg')

	@property
	def Ventilation(self) -> Quantity:
		return self._ventilation
	@Ventilation.setter
	def Ventilation(self, v: Quantity) -> None:
		# ASHRAE Fundamentals 2021, p18.15 Elevation Correction Examples
		self._ventilation = v.to('m ** 3/second')
		h = self.weather_data.altitude.m
		self._cs = 1230*(1-(h*2.25577*1e-5)) ** 5.2559
		self._cl = 3010*(1-(h*2.25577*1e-5)) ** 5.2559

	@property
	def Area(self) -> Quantity:
		return self.Width * self.Length
	#endregion
