# -*- coding: utf-8 -*-
import scrapy
import json
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active
ws.append(['Model Year', 'Make', 'Model', 'Price', 'Body Position', 'Suggested Use', 'Electric Bike Class',
           'Warranty', 'Availability', 'Motor Type', 'Motor Nominal Output', 'Motor Peak Output', 'Motor Brand',
           'Motor Torque', 'Battery Voltage', 'Battery Amp Hours', 'Battery Watt Hours', 'Battery Brand',
           'Battery Chemistry', 'Charge Time', 'Estimated Min Range', 'Display Accessories', 'Display Type',
           'Readouts', 'Drive Mode', 'Top Speed', 'Estimated Max Range', 'Total Weight', 'Battery Weight',
           'Motor Weight', 'Frame Types', 'Frame Sizes', 'Frame Material', 'Frame Colors', 'Geometry Measurements',
           'Frame Fork Details', 'Frame Rear Details', 'Attachment Points', 'Gearing Details', 'Shifter Details',
           'Cranks', 'Pedals', 'Headset', 'Stem', 'Handlebar', 'Brake Details', 'Grips', 'Saddle', 'Seat Post',
           'Seat Post Length', 'Seat Post Diameter', 'Rims', 'Spokes', 'Tire Brand', 'Wheel Sizes', 'Tire Details',
           'Tube Details', 'Accessories', 'Other'])


class BicyclesScrapeSpider(scrapy.Spider):
    name = 'bicycles_scrape'
    allowed_domains = ['electricbikereview.com/']
    start_urls = [
        'https://electricbikereview.com/wp-json/awesome-search/search/?object_class=AR_Review&group%5Btaxonomy%5D=review_cat&group%5Bterms%5D=mid-drive&group%5Bfield%5D=slug&args%5Bpost_types%5D%5Breview%5D%5Bslug%5D=review&args%5Bpost_types%5D%5Breview%5D%5Bobject_class%5D=AR_Review&args%5Bpost_types%5D%5Baccessory%5D%5Bslug%5D=accessory&args%5Bpost_types%5D%5Baccessory%5D%5Bobject_class%5D=AR_Accessory&args%5Bposts_per_page%5D=10&args%5Bobject_class%5D=AR_Review&args%5Bgroup%5D%5Btaxonomy%5D=review_cat&args%5Bgroup%5D%5Bterms%5D=mid-drive&args%5Bgroup%5D%5Bfield%5D=slug&args%5Bpagination%5D%5BresultsPerPage%5D=4&args%5Bpaged%5D=3&args%5Bpage%5D={}&pagination%5BresultsPerPage%5D=4'.format(
            page) for page in range(62)]

    def parse(self, response):
        rsp = json.loads(response.body_as_unicode())
        for row in rsp['result']:
            model_year = self.get_value_or_NA(row, 'model_year')
            make = self.get_value_or_NA(row, 'make')
            model = self.get_value_or_NA(row, 'model')
            price = self.get_value_or_NA(row, 'price')
            body_position = self.get_value_or_NA(row, 'body_position')
            suggested_use = self.get_value_or_NA(row, 'suggested_use')
            electric_bike_class = self.get_value_or_NA(row, 'electric_bike_class')
            warranty = self.get_value_or_NA(row, 'warranty')
            availability = self.get_value_or_NA(row, 'availability')
            motor_type = self.get_value_or_NA(row, 'motor_type')
            motor_nominal_output = self.get_value_or_NA(row, 'motor_nominal_output')
            motor_peak_output = self.get_value_or_NA(row, 'motor_peak_output')
            motor_brand = self.get_value_or_NA(row, 'motor_brand')
            motor_torque = self.get_value_or_NA(row, 'motor_torque')
            battery_voltage = self.get_value_or_NA(row, 'battery_voltage')
            battery_amp_hours = self.get_value_or_NA(row, 'battery_amp_hours')
            battery_watt_hours = self.get_value_or_NA(row, 'battery_watt_hours')
            battery_brand = self.get_value_or_NA(row, 'battery_brand')
            battery_chemistry = self.get_value_or_NA(row, 'battery_chemistry')
            charge_time = self.get_value_or_NA(row, 'charge_time')
            estimated_min_range = self.get_value_or_NA(row, 'estimated_min_range')
            display_accessories = self.get_value_or_NA(row, 'display_accessories')
            display_type = self.get_value_or_NA(row, 'display_type')
            readouts = self.get_value_or_NA(row, 'readouts')
            drive_mode = self.get_value_or_NA(row, 'drive_mode')
            top_speed = self.get_value_or_NA(row, 'top_speed')
            estimated_max_range = self.get_value_or_NA(row, 'estimated_max_range')
            total_weight = self.get_value_or_NA(row, 'total_weight')
            battery_weight = self.get_value_or_NA(row, 'battery_weight')
            motor_weight = self.get_value_or_NA(row, 'motor_weight')
            frame_types = self.get_value_or_NA(row, 'frame_types')
            frame_sizes = self.get_value_or_NA(row, 'frame_sizes')
            frame_material = self.get_value_or_NA(row, 'frame_material')
            frame_colors = self.get_value_or_NA(row, 'frame_colors')
            geometry_measurements = self.get_value_or_NA(row, 'geometry_measurements')
            frame_fork_details = self.get_value_or_NA(row, 'frame_fork_details')
            frame_rear_details = self.get_value_or_NA(row, 'frame_rear_details')
            attachment_points = self.get_value_or_NA(row, 'attachment_points')
            gearing_details = self.get_value_or_NA(row, 'gearing_details')
            shifter_details = self.get_value_or_NA(row, 'shifter_details')
            cranks = self.get_value_or_NA(row, 'cranks')
            pedals = self.get_value_or_NA(row, 'pedals')
            headset = self.get_value_or_NA(row, 'headset')
            stem = self.get_value_or_NA(row, 'stem')
            handlebar = self.get_value_or_NA(row, 'handlebar')
            brake_details = self.get_value_or_NA(row, 'brake_details')
            grips = self.get_value_or_NA(row, 'grips')
            saddle = self.get_value_or_NA(row, 'saddle')
            seat_post = self.get_value_or_NA(row, 'seat_post')
            seat_post_length = self.get_value_or_NA(row, 'seat_post_length')
            seat_post_diameter = self.get_value_or_NA(row, 'seat_post_diameter')
            rims = self.get_value_or_NA(row, 'rims')
            spokes = self.get_value_or_NA(row, 'spokes')
            tire_brand = self.get_value_or_NA(row, 'tire_brand')
            wheel_sizes = self.get_value_or_NA(row, 'wheel_sizes')
            tire_details = self.get_value_or_NA(row, 'tire_details')
            tube_details = self.get_value_or_NA(row, 'tube_details')
            accessories = self.get_value_or_NA(row, 'accessories')
            other = self.get_value_or_NA(row, 'other')

            ws.append([
                model_year, make, model, price, body_position, suggested_use, electric_bike_class, warranty,
                availability, motor_type, motor_nominal_output, motor_peak_output, motor_brand, motor_torque,
                battery_voltage, battery_amp_hours, battery_watt_hours, battery_brand, battery_chemistry, charge_time,
                estimated_min_range, display_accessories, display_type, readouts, drive_mode, top_speed,
                estimated_max_range, total_weight, battery_weight, motor_weight, frame_types, frame_sizes,
                frame_material, frame_colors, geometry_measurements, frame_fork_details, frame_rear_details,
                attachment_points, gearing_details, shifter_details, cranks, pedals, headset, stem, handlebar,
                brake_details, grips, saddle, seat_post, seat_post_length, seat_post_diameter, rims,
                spokes, tire_brand, wheel_sizes, tire_details, tube_details, accessories, other
            ])

        wb.save("table.xlsx")

    def get_value_or_NA(self, row: json, value: str) -> str:
        if value in row['metas']:
            return row['metas'][value]
        else:
            return 'NA'
