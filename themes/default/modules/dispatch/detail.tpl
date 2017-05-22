<!-- BEGIN: main -->
<div class="table-responsive">
	<table class="table table-striped table-bordered table-hover">
		<colgroup>
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col class="w100" />
			<col span="4" class="w150" />
		</colgroup>
		<thead>
			<tr>
				<th><strong>{LANG.dis}</strong></th>
				<th><strong>{LANG.dis_code}</strong></th>
				<th><strong>{LANG.dis_date_re}</strong></th>
				<th><strong>{LANG.dis_souce}</strong></th>
				<th><strong>{LANG.copy_count}</strong></th>
				<th><strong>{LANG.dis_person}</strong></th>
				<th><strong>{LANG.dis_to_org}</strong></th>
				<th><strong>{LANG.dis_de}</strong></th>
				<th><strong>{LANG.dis_date_iss}</strong></th>
				<th><strong>{LANG.dis_date_first}</strong></th>
				<th><strong>{LANG.dis_date_die}</strong></th>
				<th><strong>{LANG.view}</strong></th>
				<th><strong>{LANG.file}</strong></th>
				<th><strong>{LANG.dis_content}</strong></th>
				</tr>
		</thead>
		<tbody>
			<tr>
				<td><strong><a href="{TYPELINK}">{TYPENAME}</a> </strong></td>
				<td> {ROW.code} </td>
				<td> {ROW.from_time} </td>
				<td> {ROW.from_org} </td>
				<td> {ROW.copy_count} </td>
				<td> {ROW.signer} </td>
				<td> {ROW.to_org} </td>
				<td><!-- BEGIN: depid --> - {dis_de}
				<br />
				<!-- END: depid --></td>
				<td> {ROW.date_iss} </td>
				<td> {ROW.date_first} </td>
				<td> {ROW.date_die} </td>
				<td> {ROW.view} </td>
				<td>
					<!-- BEGIN: taifile -->
					<!-- BEGIN: row --><a href="{NV_BASE_SITEURL}uploads/{module}/{FILEUPLOAD}"><em class="fa fa-download">&nbsp;</em>{FILEUPLOAD} </a>
					<br />
					<!-- END: row -->
					<!-- END: taifile -->
				</td>
				<td colspan="2" style="text-align: justify;"> {ROW.content} </td>
			</tr>
		</tbody>
	</table>
</div>
<div style="clear: both;"></div>
<!-- END: main -->
<!-- BEGIN: no_permission -->

<!-- END: no_permission -->