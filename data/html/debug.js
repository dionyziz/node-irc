function expand( id ) {
		var domelement;
		
		domelement = document.getElementById( id );
		if ( domelement.style.display == 'none' ) {
				domelement.style.display = '';
		}
		else {
				domelement.style.display = 'none';
		}
}