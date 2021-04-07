import re
import argparse
import xlrd

from xlwt import Workbook
import xlwt

################################
##    PARAMETERS

min_bacs_present = 3 # minimum number of accessions that have a SNP value (only bacs)
min_accs_present = 3 # minimum number of accessions that have a SNP value (only accs)

max_error_pct = 0 # percentage max of accessions with false snp

min_SNP_acc_present = 1 # minimum acc with the same snp to be diagnostic

################################



parser = argparse.ArgumentParser(description='Extract representative SNP.')

parser.add_argument('--vcf', help='Input vcf.', required=True)
parser.add_argument('--names', help="Column file with group's name and accessions's name", required=True)
parser.add_argument('--out', help="Name of the output file", required=True)
parser.add_argument('--cor', help="Correspondence ID and name", required=False)



args = parser.parse_args()
vcf = args.vcf
names = args.names
output = args.out
cor = args.cor

print('\n\n')

print(vcf)
print(names)

cor_in = cor
vcf_in = vcf
names_in = names


###############################
# Dict correspondence

dict_cor = {}
if cor_in:
	cor_input = open(cor_in, 'r')
	for line in cor_input:
		line_split = re.split(r'\t+', line)

		dict_cor[line_split[0]] = line_split[1].rstrip("\n")

###############################
# Dict names

names_input = open(names_in,'r')
dict_names = {}
order_acc_in_file = {}
# File 'names_in' in a dictionnary. Key : Group start wirh a # (#officinarum) ; Sub-key : Name of acc belonging to the group (ex: Off_SRR6680698) ;
# Value : polymorphism data (ex :./.:24,0,0:24:0) updated each line of the vcf

for line in names_input:
	if line != '\n' :
		if line[0] == "#":
			group = line.rstrip("\n")
			dict_names[line.rstrip("\n")] = {}
			order_acc_in_file[line.rstrip("\n")] = []
		else:
			dict_names[group][line.rstrip("\n")] = ''
			order_acc_in_file[group].append(line.rstrip("\n"))


###############################
# Vcf extract

vcf_out = output + "_BFA_M" + str(min_bacs_present + min_accs_present) + "-E" + str(max_error_pct) +".vcf"
tab_out = output + "_BFA_M" + str(min_bacs_present + min_accs_present) + "-E" + str(max_error_pct) +".xls"

vcf_input = open(vcf_in,'r') 
vcf_output = open(vcf_out, 'w')

classeur = xlwt.Workbook()

pop = 1
feuille = classeur.add_sheet("SNP data_"+str(pop))


#######################################
# Define colors cells

xlwt.add_palette_colour("officinarum", 0x20)
classeur.set_colour_RGB(0x20, 2, 110, 12)

xlwt.add_palette_colour("robustum", 0x21)
classeur.set_colour_RGB(0x21, 0, 178, 238)

xlwt.add_palette_colour("spont_1", 0x22)
classeur.set_colour_RGB(0x22, 139, 54, 38)

xlwt.add_palette_colour("spont_2", 0x23)
classeur.set_colour_RGB(0x23, 193, 122, 38)

xlwt.add_palette_colour("spont_3", 0x24)
classeur.set_colour_RGB(0x24, 238, 220, 130)

xlwt.add_palette_colour("officinarum_robustum", 0x25)
classeur.set_colour_RGB(0x25, 124, 252, 0)

xlwt.add_palette_colour("all_spont", 0x26)
classeur.set_colour_RGB(0x26, 255, 64, 64)

#### NEW VERSION 3 ####

xlwt.add_palette_colour("barb", 0x27)
classeur.set_colour_RGB(0x26, 255, 64, 0)

xlwt.add_palette_colour("hyb", 0x28)
classeur.set_colour_RGB(0x26, 255, 64, 200)

xlwt.add_palette_colour("misc", 0x29)
classeur.set_colour_RGB(0x26, 255, 64, 100)

xlwt.add_palette_colour("out", 0x30)
classeur.set_colour_RGB(0x26, 255, 100, 64)

xlwt.add_palette_colour("sine", 0x31)
classeur.set_colour_RGB(0x26, 255, 100, 0)

xlwt.add_palette_colour("spont", 0x32)
classeur.set_colour_RGB(0x26, 255, 200, 64)


bacs = xlwt.easyxf('pattern: pattern solid, fore_colour white')

colors = xlwt.easyxf('pattern: pattern solid, fore_colour white')

top_border = xlwt.easyxf('borders: top THICK') 

robustum = xlwt.easyxf('pattern: pattern solid, fore_colour robustum')
spont_1 = xlwt.easyxf('pattern: pattern solid, fore_colour spont_1')
spont_2 = xlwt.easyxf('pattern: pattern solid, fore_colour spont_2')
spont_3 = xlwt.easyxf('pattern: pattern solid, fore_colour spont_3')
officinarum = xlwt.easyxf('pattern: pattern solid, fore_colour officinarum')
officinarum_robustum = xlwt.easyxf('pattern: pattern solid, fore_colour officinarum_robustum')
all_spont = xlwt.easyxf('pattern: pattern solid, fore_colour all_spont')
spont_1_spont_2 = xlwt.easyxf('pattern: pattern solid, fore_colour all_spont')
spont_1_spont_3 = xlwt.easyxf('pattern: pattern solid, fore_colour all_spont')
spont_2_spont_3 = xlwt.easyxf('pattern: pattern solid, fore_colour all_spont')
barb = xlwt.easyxf('pattern: pattern solid, fore_colour barb')
hyb = xlwt.easyxf('pattern: pattern solid, fore_colour hyb')
misc = xlwt.easyxf('pattern: pattern solid, fore_colour misc')
out = xlwt.easyxf('pattern: pattern solid, fore_colour out')
sine = xlwt.easyxf('pattern: pattern solid, fore_colour sine')
spont = xlwt.easyxf('pattern: pattern solid, fore_colour spont')

border_DASHED = xlwt.easyxf('borders: top DASHED')
robustum_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour robustum')
spont_1_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour spont_1')
spont_2_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour spont_2')
spont_3_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour spont_3')
officinarum_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour officinarum')
officinarum_robustum_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour officinarum_robustum')
all_spont_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour all_spont')
spont_1_spont_2_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour all_spont')
spont_1_spont_3_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour all_spont')
spont_2_spont_3_border_DASHED = xlwt.easyxf('borders: top DASHED; pattern: pattern solid, fore_colour all_spont')
barb_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour barb')
hyb_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour hyb')
misc_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour misc')
out_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour out')
sine_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour sine')
spont_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour spont')

border_THICK = xlwt.easyxf('borders: top THICK')
robustum_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour robustum')
spont_1_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour spont_1')
spont_2_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour spont_2')
spont_3_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour spont_3')
officinarum_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour officinarum')
officinarum_robustum_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour officinarum_robustum')
all_spont_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour all_spont')
spont_1_spont_2_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour all_spont')
spont_1_spont_3_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour all_spont')
spont_2_spont_3_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour all_spont')
barb_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour barb')
hyb_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour hyb')
misc_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour misc')
out_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour out')
sine_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour sine')
spont_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour spont')

robustum_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour robustum')
officinarum_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour officinarum')
spont_1_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour spont_1')
spont_2_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour spont_2')
spont_3_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour spont_3')
bacs_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour white')

colors_right_border = xlwt.easyxf('borders: right THIN; pattern: pattern solid, fore_colour white')
colors_border_DASHED = xlwt.easyxf('borders:  top DASHED; pattern: pattern solid, fore_colour white')
colors_border_THICK = xlwt.easyxf('borders: top THICK; pattern: pattern solid, fore_colour white')

#######################################


order_acc = {} # Key : Number corresponding to position of appearance in the vcf ; Value : Name of the accession (ex: {0: 'Sh_182G15', 1: 'Spont_SRR6680826'} )
column_tab = 0

for line in vcf_input:
	#Copy heading
	if(line[0] == '#' and line[1] == '#'):
		vcf_output.write(line)
	else:
		line_split = re.split(r'\t+', line)

		#line intialisation with accessions
		if(line_split[0] == '#CHROM'):
			cpt = 0
			for word in line_split:
				if cpt < 9 :
					cpt+=1
				else:
					if(word[-1:] == '\n') :
						order_acc[cpt-9] = word[:-1]
					else:
						order_acc[cpt-9] = word
					cpt+=1
			vcf_output.write(line)
			
			line_tab = 1

			for group in sorted(dict_names.keys(), reverse=True):
				for acc in order_acc_in_file[group] :
					feuille.write(line_tab, column_tab, str(acc), eval(group[1:]+'_right_border'))
					line_tab+=1

			column_tab +=1

			line_tab = 1

			for group in sorted(dict_names.keys(), reverse=True):
				for acc in order_acc_in_file[group] :
					if acc in dict_cor :
						feuille.write(line_tab, column_tab, dict_cor[acc], eval(group[1:]+'_right_border'))
					else:
						feuille.write(line_tab, column_tab, str(acc), eval(group[1:]+'_right_border'))
					line_tab+=1

			column_tab +=1



		##SNP study
		else :


			cpt = 0
			for word in line_split:
				if cpt > 8 : # shift until first data of type ./.:24,0,0:24:0
					for key in dict_names: #update dict_names
						if(word[-1:] == '\n') :
							if order_acc[cpt-9] in dict_names[key]:
								dict_names[key][order_acc[cpt-9]]=word[:-1]
						else:
							if order_acc[cpt-9] in dict_names[key]:
								dict_names[key][order_acc[cpt-9]]=word
				cpt+=1

			stats = {} # Key : Groupe ; Sub-key : data of type ./. ou 0/1 ; Value : number of accs of this type
			for key in dict_names:
				stats[key] = {}
				for acc in dict_names[key]:
					if dict_names[key][acc][0] != '.':
						if dict_names[key][acc][:3] in stats[key]:
							stats[key][dict_names[key][acc][:3]] +=1
						else:
							stats[key][dict_names[key][acc][:3]] = 1

			# Compute the number of accs with a data
			val_acc = 0
			val_bac = 0

			for group in stats:
				for snp in stats[group]:

					if group != '#bacs' and group != '#colors' and snp != '.' :
						val_acc += stats[group][snp]

					if group == '#bacs' and snp != '.' :
						val_bac += stats[group][snp]


			# If nb of accs i sufficient, looking for specifics SNP
			if val_bac >= min_bacs_present and val_acc >= min_accs_present:
				carac_snp = {} # Key : poly data ex 0/1 ou 0/0 ; Value : list of groups having accessions like that
				for group in stats:
					total_snp = 0
					if group != '#bacs' and group != '#colors':
						for snp in stats[group]:
							total_snp += stats[group][snp]

						for snp in stats[group]:
							if (int(stats[group][snp])/int(total_snp))*100 > max_error_pct and int(stats[group][snp]) >= min_SNP_acc_present:
								if not(snp in carac_snp):
									carac_snp[snp] = [group]
								else:
									carac_snp[snp].append(group)

				if(len(list(carac_snp.keys())) > 1): # if we have at least two different poly
					diff = 0

					carac = {} #groupe:list [values] | for groups and SNP caracterics
					for snp_1 in carac_snp:
						if len(carac_snp[snp_1]) == 1 :
							if stats[carac_snp[snp_1][0]][snp_1] >= min_SNP_acc_present :
								if not(carac_snp[snp_1][0] in carac) :
									carac[carac_snp[snp_1][0]] = [snp_1]
								else:
									carac[carac_snp[snp_1][0]].append(snp_1)

						else:
							for snp_2 in carac_snp:
								if carac_snp[snp_1] != carac_snp[snp_2]:
									if len(carac_snp[snp_1]) == 2 and '#officinarum' in carac_snp[snp_1] and '#robustum' in carac_snp[snp_1]:
										if not('#officinarum_robustum' in carac) :	
											carac['#officinarum_robustum'] = [snp_1]
										elif not(snp_1 in carac['#officinarum_robustum']):
											carac['#officinarum_robustum'].append(snp_1)
									if len(carac_snp[snp_1]) == 3 and '#spont_1' in carac_snp[snp_1] and '#spont_2' in carac_snp[snp_1] and '#spont_3' in carac_snp[snp_1]:
										if not('#all_spont' in carac) :	
											carac['#all_spont'] = [snp_1]
										elif not(snp_1 in carac['#all_spont']):
											carac['#all_spont'].append(snp_1)
									if len(carac_snp[snp_1]) == 2 and '#spont_1' in carac_snp[snp_1] and '#spont_2' in carac_snp[snp_1]:
										if not('#spont_1_spont_2' in carac) :	
											carac['#spont_1_spont_2'] = [snp_1]
										elif not(snp_1 in carac['#spont_1_spont_2']):
											carac['#spont_1_spont_2'].append(snp_1)
									if len(carac_snp[snp_1]) == 2 and '#spont_1' in carac_snp[snp_1] and '#spont_3' in carac_snp[snp_1]:
										if not('#spont_1_spont_3' in carac) :	
											carac['#spont_1_spont_3'] = [snp_1]
										elif not(snp_1 in carac['#spont_1_spont_3']):
											carac['#spont_1_spont_3'].append(snp_1)
									if len(carac_snp[snp_1]) == 2 and '#spont_2' in carac_snp[snp_1] and '#spont_3' in carac_snp[snp_1]:
										if not('#spont_2_spont_3' in carac) :	
											carac['#spont_2_spont_3'] = [snp_1]
										elif not(snp_1 in carac['#spont_2_spont_3']):
											carac['#spont_2_spont_3'].append(snp_1)

					genotype = {}

					for group in stats:
						genotype[group] = []
						for al in stats[group]:
							if al[0] not in genotype[group]:
								genotype[group].append(al[0])
							if al[2] not in genotype[group]:
								genotype[group].append(al[2])

					if len(carac):
					# verify if at least one bac is going to be colored

						color = 0
						for acc in sorted(dict_names["#bacs"].keys()):
							bool_test = 1

							snp_value = list(carac.values()) # extrait snp distinctif
							flat_list_snp_value = [item for sublist in snp_value for item in sublist]

							for el in flat_list_snp_value : #snp_value : list(carac.values()) # extrait snp distinctif

								# si l'accession est un bac, regarde si le bac est a l'origine du snp caracteristique (si oui, le colore de la meme couleur)				
								if not(str(dict_names["#bacs"][acc][0:3]) in carac_snp) and \
									not(any(dict_names["#bacs"][acc][0] in (list(set(carac_snp.keys()) - set(flat_list_snp_value)))[i] for i in range(0, len((list(set(carac_snp.keys()) - set(flat_list_snp_value))))))) and \
									dict_names["#bacs"][acc][0] in el and bool_test:
									color+=1
									bool_test=0
									


						# if not(color) or True:
						if color: 
							if column_tab == 256:
								pop+=1
								feuille = classeur.add_sheet("SNP data_"+str(pop))
								column_tab=0

							line_tab = 0
							feuille.write(line_tab, column_tab, str(line_split[1])) # ecriture position
							line_tab+=1

							
							for group in sorted(dict_names.keys(), reverse=True):

								acc_number = 0

								for acc in order_acc_in_file[group] :
									
									# to replace 0/0 by actual letter

									if dict_names[group][acc] != '.' :

										first_el = str(dict_names[group][acc][0])
										last_el = str(dict_names[group][acc][2])
										possibilities = line_split[4].split(',')

										if first_el == '.' : pass
										elif first_el == '0' : first_el = str(line_split[3])
										else : first_el = possibilities[int(first_el)-1]

										if last_el == '.' : pass
										elif last_el == '0' : last_el = str(line_split[3])
										else : last_el = possibilities[int(last_el)-1]

										snp_output = first_el + '/' + last_el

										if (group != '#bacs' or acc == 'R570_WGS' or acc == 'LaPurple_WGS' or acc == 'SES234_WGS') and snp_output[0] != '.' :

											if snp_output[0] == '.':
												sentence_split =  re.split(r':+', dict_names[group][acc])
												alig_split = re.split(r',+', sentence_split[3])

												results = list(map(int, alig_split))

												largest_integer = max(results) 
												results.remove(largest_integer)

												second_largest_integer = max(results) 
												snp_output= sentence_split[0] + ':' + str(largest_integer) + '/' + str(second_largest_integer)

											elif snp_output[0] == snp_output[2] :
												sentence_split =  re.split(r':+', dict_names[group][acc])
												alig_split = re.split(r',+', sentence_split[3])

												results = list(map(int, alig_split))

												largest_integer = max(results) 


												snp_output = first_el + '/' + last_el + ':' + str(largest_integer)

											else :

												sentence_split =  re.split(r':+', dict_names[group][acc])
												alig_split = re.split(r',+', sentence_split[3])

												results = list(map(int, alig_split))

												largest_integer = max(results) 
												results.remove(largest_integer)

												second_largest_integer = max(results) 

												pos_largest = [i for i,x in enumerate(list(map(int, alig_split))) if x == largest_integer][0]
												pos_second_largest = [i for i,x in enumerate(list(map(int, alig_split))) if x == second_largest_integer][0]


												if pos_largest < pos_second_largest:
													snp_output = first_el + '/' + last_el + ':' + str(largest_integer) + ',' + str(second_largest_integer)

													if acc == 'R570_WGS' :
														val = "{0:.1f}".format((second_largest_integer)/((largest_integer+second_largest_integer)/12))
														snp_output += ' - '+ str(12-float(val)) + '/' + val
													elif acc == 'LaPurple_WGS' or acc == 'SES234_WGS' : 
														val = "{0:.1f}".format((second_largest_integer)/((largest_integer+second_largest_integer)/8))
														snp_output += ' - '+ str(8-float(val)) + '/' + val

												else:
													snp_output = last_el + '/' +  first_el + ':' + str(largest_integer) + ',' + str(second_largest_integer)

													if acc == 'R570_WGS' :
														val = "{0:.1f}".format((second_largest_integer)/((largest_integer+second_largest_integer)/12))
														snp_output += ' - '+ str(12-float(val)) + '/' + val
													elif acc == 'LaPurple_WGS' or acc == 'SES234_WGS' : 
														val = "{0:.1f}".format((second_largest_integer)/((largest_integer+second_largest_integer)/8))
														snp_output += ' - '+ str(8-float(val)) + '/' + val



										else :
											snp_output = first_el + '/' + last_el

									else : 
										snp_output = './.'

									snp_value = list(carac.values()) # extrait snp distinctif
									flat_list_snp_value = [item for sublist in snp_value for item in sublist]

									
									if dict_names[group][acc][:3] in flat_list_snp_value and group != "#bacs" :
										bool_test = 1
										for group_bis in carac:
											if dict_names[group][acc][:3] in carac[group_bis] and bool_test:

												if not acc_number :
													feuille.write(line_tab, column_tab, str(snp_output), eval(group_bis[1:]+'_border_DASHED'))

													
												else :
													feuille.write(line_tab, column_tab, str(snp_output), eval(group_bis[1:]))

												bool_test = 0
										acc_number +=1


									else:
										if group == "#bacs" :
											bool_test = 1
											for el in flat_list_snp_value : #snp_value : list(carac.values()) # extract caracteristic SNP

												# if acc is a BAC, look if the BAC is at the origine of the caracteristique SNP (if yes, we color him with the same color)
												if not(str(dict_names[group][acc][0:3]) in carac_snp) and \
													not(any(dict_names[group][acc][0] in (list(set(carac_snp.keys()) - set(flat_list_snp_value)))[i] for i in range(0, len((list(set(carac_snp.keys()) - set(flat_list_snp_value))))))) and \
													dict_names[group][acc][0] in el and bool_test:

													for group_bis in carac:
														if any(dict_names[group][acc][0] in snp for snp in carac[group_bis]) and bool_test:
															if not acc_number :
																feuille.write(line_tab, column_tab, str(snp_output), eval(group_bis[1:]+'_border_THICK'))
															else :
																feuille.write(line_tab, column_tab, str(snp_output), eval(group_bis[1:]))

															bool_test=0

											if bool_test: #if BAC not caracteristic, we don't color it

													if not acc_number:
														feuille.write(line_tab, column_tab, str(snp_output), border_THICK)

													else :
														feuille.write(line_tab, column_tab, str(snp_output))
											acc_number+=1

										else: #if normal group not diagostic, we don't color it
											if not acc_number :
												feuille.write(line_tab, column_tab, str(snp_output), border_DASHED)

											else :
												feuille.write(line_tab, column_tab, str(snp_output))

											acc_number+=1

									line_tab+=1


							vcf_output.write(line)
							column_tab +=1


classeur.save(tab_out)



print('\n\n')
