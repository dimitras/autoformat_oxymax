# https://dl.bintray.com/oneclick/rubyinstaller/rubyinstaller-1.9.3-p551.exe
# gem install nokogiri -v 1.5.6
# gem install rubyXL -v 3.3.22
# ruby auto_xlsx_format.rb data.xlsx 8 "2015-02-24 7:00"

require 'rubyXL'
require 'time'

###input
ifile = ARGV[0] #an .xlsx file
worksheets_num = ARGV[1].to_i # number of sheets, here is 8
init_time = ARGV[2] # the date to initialize, here 2015-09-14 7:00

###init
workbook = RubyXL::Parser.parse(ifile)
# inittime = DateTime.parse("2015-09-14 7:00",'%Y-%m-%d %H:%M')
inittime = DateTime.parse(init_time,'%Y-%m-%d %H:%M')

ifile_name = File.basename(ifile)
ifile_dir = File.dirname(ifile)

# add AVG worksheets
VO2 = workbook.add_worksheet("VO2")
RER = workbook.add_worksheet("RER")
HEAT = workbook.add_worksheet("Heat")
FEED1 = workbook.add_worksheet("Food intake")
XAMB = workbook.add_worksheet("XAmb")
WHEEL = workbook.add_worksheet("Wheel running")
O2IN = workbook.add_worksheet("O2IN")
O2OUT = workbook.add_worksheet("O2OUT")
DO2 = workbook.add_worksheet("DO2")
ACCO2 = workbook.add_worksheet("ACCO2")
VCO2 = workbook.add_worksheet("VCO2")
CO2IN = workbook.add_worksheet("CO2IN")
CO2OUT = workbook.add_worksheet("CO2OUT")
DCO2 = workbook.add_worksheet("DCO2")
ACCCO2 = workbook.add_worksheet("ACCCO2")
FLOW = workbook.add_worksheet("FLOW")
FEED1ACC = workbook.add_worksheet("FEED1ACC")
DRINK1 = workbook.add_worksheet("DRINK1")
DRINK1ACC = workbook.add_worksheet("DRINK1ACC")
XTOT = workbook.add_worksheet("XTOT")
ZTOT = workbook.add_worksheet("ZTOT")
WHEELACC = workbook.add_worksheet("WHEELACC")
TEMP = workbook.add_worksheet("TEMP")
RHSAMP = workbook.add_worksheet("RHSAMP")
RHPURGE = workbook.add_worksheet("RHPURGE")
RHAMB = workbook.add_worksheet("RHAMB")
TEMPAMB = workbook.add_worksheet("TEMPAMB")
BAROPRESS = workbook.add_worksheet("BAROPRESS")
ENCLOSURETEMP = workbook.add_worksheet("ENCLOSURETEMP")
ENCLOSURESETPOINT = workbook.add_worksheet("ENCLOSURESETPOINT")	


(1..worksheets_num).each do |col|
	VO2.add_cell(0,col,"VO2")
	RER.add_cell(0,col,"RER")
	HEAT.add_cell(0,col,"HEAT")
	FEED1.add_cell(0,col,"Feed1")
	XAMB.add_cell(0,col,"XAmb")
	WHEEL.add_cell(0,col,"Wheel")
	O2IN.add_cell(0,col,"O2IN")
	O2OUT.add_cell(0,col,"O2OUT")
	DO2.add_cell(0,col,"DO2")
	ACCO2.add_cell(0,col,"ACCO2")
	VCO2.add_cell(0,col,"VCO2")
	CO2IN.add_cell(0,col,"CO2IN")
	CO2OUT.add_cell(0,col,"CO2OUT")
	DCO2.add_cell(0,col,"DCO2")
	ACCCO2.add_cell(0,col,"ACCCO2")
	FLOW.add_cell(0,col,"FLOW")
	FEED1ACC.add_cell(0,col,"FEED1ACC")
	DRINK1.add_cell(0,col,"DRINK1")
	DRINK1ACC.add_cell(0,col,"DRINK1ACC")
	XTOT.add_cell(0,col,"XTOT")
	ZTOT.add_cell(0,col,"ZTOT")
	WHEELACC.add_cell(0,col,"WHEELACC")
	TEMP.add_cell(0,col,"TEMP")
	RHSAMP.add_cell(0,col,"RHSAMP")
	RHPURGE.add_cell(0,col,"RHPURGE")
	RHAMB.add_cell(0,col,"RHAMB")
	TEMPAMB.add_cell(0,col,"TEMPAMB")
	BAROPRESS.add_cell(0,col,"BAROPRESS")
	ENCLOSURETEMP.add_cell(0,col,"ENCLOSURETEMP")
	ENCLOSURESETPOINT.add_cell(0,col,"ENCLOSURESETPOINT")
end

###read the sheets
sheet_names = {}
data_array = []
(0...worksheets_num).each do |ws_num|
	sheet_names[workbook[ws_num].sheet_name] = ws_num+1
	worksheet = workbook[ws_num]

	# worksheet.insert_column(35)

	in_data_section = false
	row_num = 0
	vo2_ranges = Hash.new { |h,k| h[k] = [] }
	rer_ranges = Hash.new { |h,k| h[k] = [] }
	heat_ranges = Hash.new { |h,k| h[k] = [] }
	feed1_ranges = Hash.new { |h,k| h[k] = [] }
	xamb_ranges = Hash.new { |h,k| h[k] = [] }
	wheel_ranges = Hash.new { |h,k| h[k] = [] }
	o2in_ranges = Hash.new { |h,k| h[k] = [] }
	o2out_ranges = Hash.new { |h,k| h[k] = [] }
	do2_ranges = Hash.new { |h,k| h[k] = [] }
	acco2_ranges = Hash.new { |h,k| h[k] = [] }
	vco2_ranges = Hash.new { |h,k| h[k] = [] }
	co2in_ranges = Hash.new { |h,k| h[k] = [] }
	co2out_ranges = Hash.new { |h,k| h[k] = [] }
	dco2_ranges = Hash.new { |h,k| h[k] = [] }
	accco2_ranges = Hash.new { |h,k| h[k] = [] }
	flow_ranges = Hash.new { |h,k| h[k] = [] }
	feed1acc_ranges = Hash.new { |h,k| h[k] = [] }
	drink1_ranges = Hash.new { |h,k| h[k] = [] }
	drink1acc_ranges = Hash.new { |h,k| h[k] = [] }
	xtot_ranges = Hash.new { |h,k| h[k] = [] }
	ztot_ranges = Hash.new { |h,k| h[k] = [] }
	wheelacc_ranges = Hash.new { |h,k| h[k] = [] }
	temp_ranges = Hash.new { |h,k| h[k] = [] }
	rhsamp_ranges = Hash.new { |h,k| h[k] = [] }
	rhpurge_ranges = Hash.new { |h,k| h[k] = [] }
	rhamb_ranges = Hash.new { |h,k| h[k] = [] }
	tempamb_ranges = Hash.new { |h,k| h[k] = [] }
	baropress_ranges = Hash.new { |h,k| h[k] = [] }
	enclosuretemp_ranges = Hash.new { |h,k| h[k] = [] }
	enclosuresetpoint_ranges = Hash.new { |h,k| h[k] = [] }

	worksheet.each do |row|
		if !row.nil?
			col_num = 0

			if row[0].value.to_i == 1
		    	in_data_section = true
		    end

			row && row.cells.each do |cell|
				val = cell && cell.value

				if row[0].value.nil?
			    	in_data_section = false
			    end

			    if in_data_section == true
			    	# if row.nil? || row[0].value.nil? || row[0].value == ":EVENTS"
			    	# 	in_data_section = false
			    	# end
			    	if col_num == 2
				    	# worksheet[row_num][col_num].change_contents(val, ((val-inittime)*24).to_i)
				    	# puts val
				    	val = ((val-inittime)*24).to_i
				    	vo2_ranges[val] << row[3].value.to_f if !row[3].nil?
				    	o2in_ranges[val] << row[4].value.to_f if !row[4].nil?
				    	o2out_ranges[val] << row[5].value.to_f if !row[5].nil?
				    	do2_ranges[val] << row[6].value.to_f if !row[6].nil?
				    	acco2_ranges[val] << row[7].value.to_f if !row[7].nil?
				    	vco2_ranges[val] << row[8].value.to_f if !row[8].nil?
				    	co2in_ranges[val] << row[9].value.to_f if !row[9].nil?
				    	co2out_ranges[val] << row[10].value.to_f if !row[10].nil?
				    	dco2_ranges[val] << row[11].value.to_f if !row[11].nil?
				    	accco2_ranges[val] << row[12].value.to_f if !row[12].nil?
				    	rer_ranges[val] << row[13].value.to_f if !row[13].nil?
						heat_ranges[val] << row[14].value.to_f if !row[14].nil?
						flow_ranges[val] << row[15].value.to_f if !row[15].nil?
						feed1_ranges[val] << row[17].value.to_f if !row[17].nil?
						feed1acc_ranges[val] << row[18].value.to_f if !row[18].nil?
						drink1_ranges[val] << row[19].value.to_f if !row[19].nil?
						drink1acc_ranges[val] << row[20].value.to_f if !row[20].nil?						
						xtot_ranges[val] << row[21].value.to_f if !row[21].nil?
						xamb_ranges[val] << row[22].value.to_f if !row[22].nil?
						ztot_ranges[val] << row[23].value.to_f if !row[23].nil?
						wheel_ranges[val] << row[24].value.to_f if !row[24].nil?
						wheelacc_ranges[val] << row[25].value.to_f if !row[25].nil?
						temp_ranges[val] << row[26].value.to_f if !row[26].nil?
						rhsamp_ranges[val] << row[27].value.to_f if !row[27].nil?
						rhpurge_ranges[val] << row[28].value.to_f if !row[28].nil?
						rhamb_ranges[val] << row[29].value.to_f if !row[29].nil?
						tempamb_ranges[val] << row[30].value.to_f if !row[30].nil?
						baropress_ranges[val] << row[31].value.to_f if !row[31].nil?
						enclosuretemp_ranges[val] << row[33].value.to_f if !row[33].nil?
						enclosuresetpoint_ranges[val] << row[34].value.to_f if !row[34].nil?
				    end
				end
				col_num += 1
			end
		elsif row.nil?
			in_data_section = false
		end
		row_num += 1
	end

	# create the VO2 AVG table
	VO2.add_cell(0,ws_num+1,"VO2 #{ws_num+1}")
	i=0
	vo2_ranges.each_key do |range|
		i += 1
		VO2.add_cell(i,0,range)
		# puts "sheet num: #{ws_num} and i= #{i}"
		# puts "ARRAY for #{range}: #{vo2_ranges[range].inspect}"
		average_for_range = vo2_ranges[range].inject{|sum, el| sum + el}.to_f/vo2_ranges[range].size
		# puts "AVG: #{average_for_range}"
		VO2.add_cell(i,ws_num+1,average_for_range)
	end
	
	# create the RER AVG table
	RER.add_cell(0,ws_num+1,"RER #{ws_num+1}")
	i2=0
	rer_ranges.each_key do |range|
		i2 += 1
		RER.add_cell(i2,0,range)
		average_for_range = rer_ranges[range].inject{|sum, el| sum + el}.to_f/rer_ranges[range].size
		RER.add_cell(i2,ws_num+1,average_for_range)
	end

	# create the HEAT AVG table
	HEAT.add_cell(0,ws_num+1,"HEAT #{ws_num+1}")
	i3=0
	heat_ranges.each_key do |range|
		i3 += 1
		HEAT.add_cell(i3,0,range)
		average_for_range = heat_ranges[range].inject{|sum, el| sum + el}.to_f/heat_ranges[range].size
		HEAT.add_cell(i3,ws_num+1,average_for_range)
	end

	# create the FEED1 AVG table
	FEED1.add_cell(0,ws_num+1,"Feed1 #{ws_num+1}")
	i4=0
	feed1_ranges.each_key do |range|
		i4 += 1
		FEED1.add_cell(i4,0,range)
		average_for_range = feed1_ranges[range].inject{|sum, el| sum + el}.to_f/feed1_ranges[range].size
		FEED1.add_cell(i4,ws_num+1,average_for_range)
	end

	# create the XAMB AVG table
	XAMB.add_cell(0,ws_num+1,"XAmb #{ws_num+1}")
	i5=0
	xamb_ranges.each_key do |range|
		i5 += 1
		XAMB.add_cell(i5,0,range)
		average_for_range = xamb_ranges[range].inject{|sum, el| sum + el}.to_f/xamb_ranges[range].size
		XAMB.add_cell(i5,ws_num+1,average_for_range)
	end
	
	# create the Wheel AVG table
	WHEEL.add_cell(0,ws_num+1,"Wheel #{ws_num+1}")
	i6=0
	wheel_ranges.each_key do |range|
		i6 += 1
		WHEEL.add_cell(i6,0,range)
		average_for_range = wheel_ranges[range].inject{|sum, el| sum + el}.to_f/wheel_ranges[range].size
		WHEEL.add_cell(i6,ws_num+1,average_for_range)
	end

	# create the o2in AVG table
	O2IN.add_cell(0,ws_num+1,"O2IN #{ws_num+1}")
	i7=0
	o2in_ranges.each_key do |range|
		i7 += 1
		O2IN.add_cell(i7,0,range)
		average_for_range = o2in_ranges[range].inject{|sum, el| sum + el}.to_f/o2in_ranges[range].size
		O2IN.add_cell(i7,ws_num+1,average_for_range)
	end

	# create the O2OUT AVG table
	O2OUT.add_cell(0,ws_num+1,"O2OUT #{ws_num+1}")
	i8=0
	o2out_ranges.each_key do |range|
		i8 += 1
		O2OUT.add_cell(i8,0,range)
		average_for_range = o2out_ranges[range].inject{|sum, el| sum + el}.to_f/o2out_ranges[range].size
		O2OUT.add_cell(i8,ws_num+1,average_for_range)
	end

	# create the DO2 AVG table
	DO2.add_cell(0,ws_num+1,"DO2 #{ws_num+1}")
	i9=0
	do2_ranges.each_key do |range|
		i9 += 1
		DO2.add_cell(i9,0,range)
		average_for_range = do2_ranges[range].inject{|sum, el| sum + el}.to_f/do2_ranges[range].size
		DO2.add_cell(i9,ws_num+1,average_for_range)
	end

	# create the ACCO2 AVG table
	ACCO2.add_cell(0,ws_num+1,"ACCO2 #{ws_num+1}")
	i10=0
	acco2_ranges.each_key do |range|
		i10 += 1
		ACCO2.add_cell(i10,0,range)
		average_for_range = acco2_ranges[range].inject{|sum, el| sum + el}.to_f/acco2_ranges[range].size
		ACCO2.add_cell(i10,ws_num+1,average_for_range)
	end

	# create the VCO2 AVG table
	VCO2.add_cell(0,ws_num+1,"VCO2 #{ws_num+1}")
	i11=0
	vco2_ranges.each_key do |range|
		i11 += 1
		VCO2.add_cell(i11,0,range)
		average_for_range = vco2_ranges[range].inject{|sum, el| sum + el}.to_f/vco2_ranges[range].size
		VCO2.add_cell(i11,ws_num+1,average_for_range)
	end

	# create the CO2IN AVG table
	CO2IN.add_cell(0,ws_num+1,"CO2IN #{ws_num+1}")
	i12=0
	co2in_ranges.each_key do |range|
		i12 += 1
		CO2IN.add_cell(i12,0,range)
		average_for_range = co2in_ranges[range].inject{|sum, el| sum + el}.to_f/co2in_ranges[range].size
		CO2IN.add_cell(i12,ws_num+1,average_for_range)
	end

	# create the CO2OUT AVG table
	CO2OUT.add_cell(0,ws_num+1,"CO2OUT #{ws_num+1}")
	i13=0
	co2out_ranges.each_key do |range|
		i13 += 1
		CO2OUT.add_cell(i13,0,range)
		average_for_range = co2out_ranges[range].inject{|sum, el| sum + el}.to_f/co2out_ranges[range].size
		CO2OUT.add_cell(i13,ws_num+1,average_for_range)
	end

	# create the DCO2 AVG table
	DCO2.add_cell(0,ws_num+1,"DCO2 #{ws_num+1}")
	i14=0
	dco2_ranges.each_key do |range|
		i14 += 1
		DCO2.add_cell(i14,0,range)
		average_for_range = dco2_ranges[range].inject{|sum, el| sum + el}.to_f/dco2_ranges[range].size
		DCO2.add_cell(i14,ws_num+1,average_for_range)
	end

	# create the ACCCO2 AVG table
	ACCCO2.add_cell(0,ws_num+1,"ACCCO2 #{ws_num+1}")
	i15=0
	accco2_ranges.each_key do |range|
		i15 += 1
		ACCCO2.add_cell(i15,0,range)
		average_for_range = accco2_ranges[range].inject{|sum, el| sum + el}.to_f/accco2_ranges[range].size
		ACCCO2.add_cell(i15,ws_num+1,average_for_range)
	end

	# create the FLOW AVG table
	FLOW.add_cell(0,ws_num+1,"FLOW #{ws_num+1}")
	i16=0
	flow_ranges.each_key do |range|
		i16 += 1
		FLOW.add_cell(i16,0,range)
		average_for_range = flow_ranges[range].inject{|sum, el| sum + el}.to_f/flow_ranges[range].size
		FLOW.add_cell(i16,ws_num+1,average_for_range)
	end

	# create the FEED1ACC AVG table
	FEED1ACC.add_cell(0,ws_num+1,"FEED1ACC #{ws_num+1}")
	i17=0
	feed1acc_ranges.each_key do |range|
		i17 += 1
		FEED1ACC.add_cell(i17,0,range)
		average_for_range = feed1acc_ranges[range].inject{|sum, el| sum + el}.to_f/feed1acc_ranges[range].size
		FEED1ACC.add_cell(i17,ws_num+1,average_for_range)
	end

	# create the DRINK1 AVG table
	DRINK1.add_cell(0,ws_num+1,"DRINK1 #{ws_num+1}")
	i18=0
	drink1_ranges.each_key do |range|
		i18 += 1
		DRINK1.add_cell(i18,0,range)
		average_for_range = drink1_ranges[range].inject{|sum, el| sum + el}.to_f/drink1_ranges[range].size
		DRINK1.add_cell(i18,ws_num+1,average_for_range)
	end

	# create the DRINK1ACC AVG table
	DRINK1ACC.add_cell(0,ws_num+1,"DRINK1ACC #{ws_num+1}")
	i19=0
	drink1acc_ranges.each_key do |range|
		i19 += 1
		DRINK1ACC.add_cell(i19,0,range)
		average_for_range = drink1acc_ranges[range].inject{|sum, el| sum + el}.to_f/drink1acc_ranges[range].size
		DRINK1ACC.add_cell(i19,ws_num+1,average_for_range)
	end

	# create the XTOT AVG table
	XTOT.add_cell(0,ws_num+1,"XTOT #{ws_num+1}")
	i20=0
	xtot_ranges.each_key do |range|
		i20 += 1
		XTOT.add_cell(i20,0,range)
		average_for_range = xtot_ranges[range].inject{|sum, el| sum + el}.to_f/xtot_ranges[range].size
		XTOT.add_cell(i20,ws_num+1,average_for_range)
	end

	# create the ZTOT AVG table
	ZTOT.add_cell(0,ws_num+1,"ZTOT #{ws_num+1}")
	i21=0
	ztot_ranges.each_key do |range|
		i21 += 1
		ZTOT.add_cell(i21,0,range)
		average_for_range = ztot_ranges[range].inject{|sum, el| sum + el}.to_f/ztot_ranges[range].size
		ZTOT.add_cell(i21,ws_num+1,average_for_range)
	end

	# create the WHEELACC AVG table
	WHEELACC.add_cell(0,ws_num+1,"WHEELACC #{ws_num+1}")
	i22=0
	wheelacc_ranges.each_key do |range|
		i22 += 1
		WHEELACC.add_cell(i22,0,range)
		average_for_range = wheelacc_ranges[range].inject{|sum, el| sum + el}.to_f/wheelacc_ranges[range].size
		WHEELACC.add_cell(i22,ws_num+1,average_for_range)
	end

	# create the TEMP AVG table
	TEMP.add_cell(0,ws_num+1,"TEMP #{ws_num+1}")
	i23=0
	temp_ranges.each_key do |range|
		i23 += 1
		TEMP.add_cell(i23,0,range)
		average_for_range = temp_ranges[range].inject{|sum, el| sum + el}.to_f/temp_ranges[range].size
		TEMP.add_cell(i23,ws_num+1,average_for_range)
	end

	# create the RHSAMP AVG table
	RHSAMP.add_cell(0,ws_num+1,"RHSAMP #{ws_num+1}")
	i24=0
	rhsamp_ranges.each_key do |range|
		i24 += 1
		RHSAMP.add_cell(i24,0,range)
		average_for_range = rhsamp_ranges[range].inject{|sum, el| sum + el}.to_f/rhsamp_ranges[range].size
		RHSAMP.add_cell(i24,ws_num+1,average_for_range)
	end

	# create the RHPURGE AVG table
	RHPURGE.add_cell(0,ws_num+1,"RHPURGE #{ws_num+1}")
	i25=0
	rhpurge_ranges.each_key do |range|
		i25 += 1
		RHPURGE.add_cell(i25,0,range)
		average_for_range = rhpurge_ranges[range].inject{|sum, el| sum + el}.to_f/rhpurge_ranges[range].size
		RHPURGE.add_cell(i25,ws_num+1,average_for_range)
	end

	# create the RHAMB AVG table
	RHAMB.add_cell(0,ws_num+1,"RHAMB #{ws_num+1}")
	i26=0
	rhamb_ranges.each_key do |range|
		i26 += 1
		RHAMB.add_cell(i26,0,range)
		average_for_range = rhamb_ranges[range].inject{|sum, el| sum + el}.to_f/rhamb_ranges[range].size
		RHAMB.add_cell(i26,ws_num+1,average_for_range)
	end

	# create the TEMPAMB AVG table
	TEMPAMB.add_cell(0,ws_num+1,"TEMPAMB #{ws_num+1}")
	i27=0
	tempamb_ranges.each_key do |range|
		i27 += 1
		TEMPAMB.add_cell(i27,0,range)
		average_for_range = tempamb_ranges[range].inject{|sum, el| sum + el}.to_f/tempamb_ranges[range].size
		TEMPAMB.add_cell(i27,ws_num+1,average_for_range)
	end

	# create the BAROPRESS AVG table
	BAROPRESS.add_cell(0,ws_num+1,"BAROPRESS #{ws_num+1}")
	i28=0
	baropress_ranges.each_key do |range|
		i28 += 1
		BAROPRESS.add_cell(i28,0,range)
		average_for_range = baropress_ranges[range].inject{|sum, el| sum + el}.to_f/baropress_ranges[range].size
		BAROPRESS.add_cell(i28,ws_num+1,average_for_range)
	end

	# create the ENCLOSURETEMP AVG table
	ENCLOSURETEMP.add_cell(0,ws_num+1,"ENCLOSURETEMP #{ws_num+1}")
	i29=0
	enclosuretemp_ranges.each_key do |range|
		i29 += 1
		ENCLOSURETEMP.add_cell(i29,0,range)
		average_for_range = enclosuretemp_ranges[range].inject{|sum, el| sum + el}.to_f/enclosuretemp_ranges[range].size
		ENCLOSURETEMP.add_cell(i29,ws_num+1,average_for_range)
	end

	# create the ENCLOSURESETPOINT AVG table
	ENCLOSURESETPOINT.add_cell(0,ws_num+1,"ENCLOSURESETPOINT #{ws_num+1}")
	i30=0
	enclosuresetpoint_ranges.each_key do |range|
		i30 += 1
		ENCLOSURESETPOINT.add_cell(i30,0,range)
		average_for_range = enclosuresetpoint_ranges[range].inject{|sum, el| sum + el}.to_f/enclosuresetpoint_ranges[range].size
		ENCLOSURESETPOINT.add_cell(i30,ws_num+1,average_for_range)
	end
end

workbook.write("#{ifile_dir}/modified_#{ifile_name}")
puts "modified_#{ifile_name} is generated"
