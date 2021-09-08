document.addEventListener('DOMContentLoaded', function eventHandler(event) {
		alert("event ",event.type)
		var sel = document.getElementById('options')
		//var inputVal = getSelectedOption(sel)
		const selection = 'options';
		const request = new XMLHttpRequest();
		request.open('POST', '/searchinv');
		
		// Callback function for when request is completed
		request.onload = () =>{
			//const data = JSON.parse(request.responseText)
			const indata = request.responseText
			var data = JSON.parse(indata);
			alert(data)
			var _inv = JSON.parse(data.inv_list)
			alert(data.success)
			alert(_book)
			// Update the result div
			if (data.success) {
				// Load the entire inventory
				for (active_book in inv_list) { 
					var a1 = document.createElement('a');
					a1.href = url_for('item', nv_id=inv.id);
					var node1 = document.createTextNode(inv.id);
					a1.appendChild(node1);
					var element1 = document.getElementById("id");
					element1.appendChild(a1);
					
					var a2 = document.createElement('a');
					a2.href = url_for('item', inv_id=inv.id);
					var node2 = document.createTextNode(inv.shelf);
					a2.appendChild(node2);
					var element2 = document.getElementById("shelf");
					element2.appendChild(a2);
					
					var a2 = document.createElement('a');
					a2.href = url_for('item', inv_id=inv.id);
					var node2 = document.createTextNode(inv.description);
					a2.appendChild(node2);
					var element2 = document.getElementById("desc");
					element2.appendChild(a2);
					
					var a3 = document.createElement('a');
					a3.href = url_for('item', inv_id=inv.id);
					var node3 = document.createTextNode(inv.category);
					a3.appendChild(node3);
					var element3 = document.getElementById("cat");
					element3.appendChild(a3);
					
					var a4 = document.createElement('a');
					a4.href = url_for('item', inv_id=inv.id);
					var node4 = document.createTextNode(inv.Model);
					a4.appendChild(node4);
					var element4 = document.getElementById("model");
					element4.appendChild(a4);	

					var a5 = document.createElement('a');
					a5.href = url_for('item', inv_id=inv.id);
					var node5 = document.createTextNode(inv.serial_number);
					a5.appendChild(node5);
					var element5 = document.getElementById("sn");
					element5.appendChild(a5);

					var a6 = document.createElement('a');
					a6.href = url_for('item', inv_id=inv.id);
					var node6 = document.createTextNode(inv.status);
					a6.appendChild(node6);
					var element6 = document.getElementById("status");
					element6.appendChild(a6);

					var a7 = document.createElement('a');
					a7.href = url_for('item', inv_id=inv.id);
					var node7 = document.createTextNode(inv.location);
					a7.appendChild(node7);
					var element7 = document.getElementById("loc");
					element7.appendChild(a7);

					var a8 = document.createElement('a');
					a8.href = url_for('item', inv_id=inv.id);
					var node8 = document.createTextNode(inv.quantity);
					a8.appendChild(node8);
					var element8 = document.getElementById("quant");
					element8.appendChild(a8);	

					var a9 = document.createElement('a');
					a9.href = url_for('item', inv_id=inv.id);
					var node9 = document.createTextNode(inv.active);
					a9.appendChild(node9);
					var element9 = document.getElementById("active");
					element9.appendChild(a9);		
				};
			};
		};
		// Add data to send with request
		const data = new FormData();
		inputVal = ""
		data.append("inputVal", inputVal);
		data.append('selection', selection);
		// Send request
		request.send(data);
		return false;
	
});


function getSelectedOption(sel) {
	var opt;
	for ( var i = 0, len = sel.options.length; i < len; i++ ) {
		opt = sel.options[i];
		if ( opt.selected === true ) {
			break;
		}
	}
	return opt;
}


	