<template>
	<div class="page">
		<table>
			<tbody>
				<tr style="height: 1cm">
					<td colspan="8" class="font-bold text-white bg-gray text-center">ទម្រង់បែបបទសេវាកម្មធនាគារអនឡាញ /
						Digital
						Banking Service
						Application Form</td>
				</tr>
				<tr>
					<td colspan="8" class="font-bold text-white bg-gray">សម្រាប់ធនាគារប៉ុណ្ណោះ / FOR BANK USE ONLY</td>
				</tr>
				<tr style="height: 1.5cm">
					<td colspan="8">&nbsp;</td>
				</tr>

				<!-- ព័ត៌មានអតិថិជន / Customer Information -->
				<tr>
					<td colspan="8" class="font-bold text-white bg-gray">ព័ត៌មានអតិថិជន / Customer Information</td>
				</tr>
				<tr>
					<td colspan="1">ឈ្មោះ / Name</td>
					<td colspan="2" class="bg-light-gray cus-info">{{
							props.reportDatas.name || '&nbsp;'
					}}</td>

					<td colspan="1">ថ្ងៃខែឆ្នាំកំណើត Date of Birth</td>
					<td colspan="3" class="bg-light-gray cus-info">{{ props.reportDatas.dob || '&nbsp;' }}</td>
				</tr>
				<!-- alskdjflasdfjlksfd -->
				<tr>
					<td colspan="1">សញ្ជាតិ / Nationality</td>
					<td colspan="2" class="bg-light-gray cus-info">{{
							props.reportDatas.nationality || '&nbsp;'
					}}</td>
					<td colspan="2" style="width: 2cm; max-width: 2cm;">លេខអត្តសញ្ញាណប័ណ្ណ លិខិតឆ្លងដែន ឬ លេខចុះឈ្មោះ
						<br>
						National ID / Passport / Registration Number</td>
					<td colspan="1" class="bg-light-gray cus-info">{{ props.reportDatas.national_id || '&nbsp;' }}</td>
				</tr>
				<!-- asdffasfd -->
				<tr>
					<td rowspan="2">លេខទំនាក់ទំនង / Contact</td>
					<td>លេខទូរស័ព្ទផ្ទះ / Home</td>
					<td class="bg-light-gray cus-info">{{ props.reportDatas.home_tel ||
							'&nbsp;'
					}}</td>
					<td>ការិយាល័យ / Office</td>
					<td colspan="3" class="bg-light-gray cus-info">{{ props.reportDatas.office || '&nbsp;' }}</td>
				</tr>
				<tr>
					<td>លេខទូរស័ព្ទ / Mobile</td>
					<td class="bg-light-gray cus-info">{{ props.reportDatas.mobile_tel
							|| '&nbsp;'
					}}</td>
					<td>អ៊ីម៉ែល / Email</td>
					<td colspan="3" class="bg-light-gray cus-info">{{ props.reportDatas.email || '&nbsp;' }}</td>
				</tr>
				<tr>
					<td colspan="8">
						<span style="width: 5.8cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.mobile_banking ? true : false" />
							ធនាគារចល័ត / Mobile Banking
						</span>
						<span>
							<input type="checkbox" :checked="props.reportDatas.internet_banking ? true : false" />
							ធនាគារអ៊ីនធឺណែត / Internet Banking
						</span>
					</td>
				</tr>
				<tr>
					<td colspan="8" style="border-bottom: 0;">
						<span style="width: 5.8cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.subscribe ? true : false" />
							ភ្ជាប់ / Subscribe
						</span>
						<span style="width: 3.5cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.lock ? true : false" />
							បិទ / Lock
						</span>
						<span style="width: 5.7cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.reset_password ? true : false" />
							កែពាក្យសម្ងាត់ / Reset Password
						</span>
						<span style="width: 3cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.amend ? true : false" />
							កែប្រែ / Amend
						</span>
					</td>
				</tr>
				<tr>
					<td colspan="8" style="border-top: 0;">
						<span style="width: 5.8cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.unsubscribe ? true : false" />
							ផ្តាច់ / Unsubscribe
						</span>
						<span style="width: 3.5cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.unlock ? true : false" />
							ដោះចេញ / Unlock
						</span>
						<span style="width: 5.7cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.re_active ? true : false" />
							ដំណើរការឡើងវិញ / Re-active
						</span>
					</td>
				</tr>

				<!-- ធនាគារអ៊ីនធឺណែតប៉ុណ្ណោះ / Internet Banking only -->
				<tr>
					<td colspan="8" class="font-bold text-white bg-gray">ធនាគារអ៊ីនធឺណែតប៉ុណ្ណោះ / Internet Banking only
					</td>
				</tr>
				<tr>
					<td colspan="2">
						ពិនិត្យគណនី / Inquiry Account
					</td>
					<td colspan="6">
						<span style="width: 5cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.allAccount ? true : false" />
							គ្រប់គណនី / All accounts
						</span>
						<span style="width: 5.7cm; display: inline-block">
							<input type="checkbox" :checked="props.reportDatas.re_active ? true : false" />
							គណនីកំណត់ប៉ុណ្ណោះ / Only Designated Account
						</span>
					</td>
				</tr>
				<tr>
					<td rowspan="2">
						កម្រិតនៃការផ្ទេរ <br> Transfer Limit
					</td>
					<td rowspan="2">
						កម្រិតប្រចាំថ្ងៃ <br> Daily Limit
					</td>
					<td colspan="2" class="bg-light-gray">
						6,500.00 ដុល្លារ / USD
					</td>
					<td rowspan="2" colspan="2">
						ទំហំទឹកប្រាក់ក្នុងមួយប្រតិបត្តិការ <br> Transaction Limit
					</td>
					<td colspan="2" class="bg-light-gray">
						5,500.00 ដុល្លារ / USD
					</td>
				</tr>
				<tr>
					<td colspan="2" class="bg-light-gray">
						26,000,000 រៀល / Riel
					</td>
					<td colspan="2" class="bg-light-gray">
						22,000,000 រៀល / Riel
					</td>
				</tr>
			</tbody>
		</table>
	</div>
</template>

<script setup lang="ts">
type PropDatas = {
	[key: string]: any;
}
const props = defineProps<{
	reportDatas: PropDatas
}>()
</script>

<style scoped>
@import url('https://fonts.googleapis.com/css2?family=Battambang:wght@400;700;900&display=swap');

.page {
	background: white;
	display: block;
	margin: 0 auto;
	padding: 0.5cm;
	box-shadow: 0 0 0.2cm rgb(0 0 0 / 20%);
	width: 21cm;
	height: 29.7cm;
	color: black;
}

table {
	font-family: 'Battambang', cursive;
	border-collapse: collapse;
	border-spacing: 0;
	border-top-width: 1px;
	border-right-width: 0px;
	border-bottom-width: 0px;
	border-left-width: 1px;
	border-color: black;
	border-style: solid;
	border-collapse: collapse;
	width: 100%;
}

td {
	border-top-width: 0px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 0px;
	border-color: black;
	border-style: solid;
	white-space: nowrap;
	font-size: 12px;
}

.font-bold {
	font-weight: 700;
}

.font-heavy {
	font-weight: 900;
}

.bg-gray {
	background-color: #808080;
}

.bg-light-gray {
	background-color: #cccccc75;
}

.text-white {
	color: white;
}

.text-center {
	text-align: center;
}

.cus-info {
	font-weight: 700;
	text-align: center;
	font-size: 14px;
	text-transform: uppercase;
}

input[type="checkbox"] {
	vertical-align: middle;
	accent-color: white;
}
</style>