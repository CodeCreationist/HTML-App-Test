//handles the ability to scroll if any part of the function fails in
// assigning the values it keeps running regardless
function scrollandmove()
{
	//fillp4f
	if ($(".is-pop4form-visible")[0]){
		var elmnt = document.getElementById("fillp4f");
		var x = elmnt.scrollLeft;
		var y = elmnt.scrollTop;
	} 
	else{
		var elmntt = document.getElementById("fillp4t");
		var x = elmntt.scrollLeft;
		var y = elmntt.scrollTop;
	}
	try{
		var close = document.getElementById("close");
		close.style.marginRight = "-"+x+"px";
		close.style.marginTop = ""+y+"px";
		var addsearch = document.getElementById("addsearch");
		addsearch.style.marginLeft = ""+x+"px";
		addsearch.style.marginTop = ""+y+"px";
		var moveexport = document.getElementById("moveexport");
		moveexport.style.marginRight = "-"+x+"px";
		moveexport.style.marginBottom = "-"+y+"px";
	}catch(e){}
}

/*
Handles Error messages that have been made through the try and catch method of
 the majority of javascript functions.
Adds a line to a error log file of the original error and what was being attempted.

*/
function ErrorToFile(e, e1, e2, e3, e4) {
	try{
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var s = fso.OpenTextFile( "ErrorLog/ErrorLog.txt", 8, false, 0);
    s.writeline("")
	s.writeline(Date());
    s.writeline(e);
	if(e1 != null)
		s.writeline(e1);
	if(e2 != null)
		s.writeline(e2);
	if(e3 != null)
		s.writeline(e3);
	if(e4 != null)
		s.writeline(e4);
	var WinNetwork = new ActiveXObject("WScript.Network");
	s.writeline("End of Error for USER : "+WinNetwork.UserName);
    s.Close();
	}catch(e){}
}
/*
Handles the menu popups
on the various pages of the HTA

*/
(function($) {
	try{

	"use strict";
	// Admin Menu.
		var	$window = $(window),
			$body = $('body'),
			$header = $('#header'),
			$banner = $('#banner');

		// Disable animations/transitions until the page has loaded.
			$body.addClass('is-loading');

			$window.on('load', function() {
				window.setTimeout(function() {
					$body.removeClass('is-loading');
				}, 100);
			});

		// Fix: Placeholder polyfill.
			$('form').placeholder();

		// Header.
			if (skel.vars.IEVersion < 9)
				$header.removeClass('alt');

			if ($banner.length > 0
			&&	$header.hasClass('alt')) {

				$window.on('resize', function() { $window.trigger('scroll'); });

				$banner.scrollex({
					bottom:		$header.outerHeight(),
					terminate:	function() { $header.removeClass('alt'); },
					enter:		function() { $header.addClass('alt'); },
					leave:		function() { $header.removeClass('alt'); }
				});

			}
	//pop4table function
$(function() {
		
			var $pop4table = $('#pop4table');

			$pop4table._locked = false;

			$pop4table._lock = function() {

				if ($pop4table._locked)
					return false;

				$pop4table._locked = true;

				window.setTimeout(function() {
					$pop4table._locked = false;
				}, 50);

				return true;

			};

			$pop4table._show = function() {

				if ($pop4table._lock())
					$body.addClass('is-pop4table-visible');

			};

			$pop4table._hide = function() {

				if ($pop4table._lock())
					$body.removeClass('is-pop4table-visible');

			};

			$pop4table._toggle = function() {

				if ($pop4table._lock())
					$body.toggleClass('is-pop4table-visible');

			};

			$pop4table
				$pop4table.appendTo($body)
				.on('click', function(event) {

					event.stopPropagation();

					// Hide.
						$pop4table._hide();

				})
				.find('.inner')
					.on('click', '.close', function(event) {

						event.preventDefault();
						event.stopPropagation();
						event.stopImmediatePropagation();

						// Hide.
							$pop4table._hide();

					})
					.on('click', function(event) {
						event.stopPropagation();
					})
					.on('click', 'a', function(event) {

						var href = $(this).attr('href');

						event.preventDefault();
						event.stopPropagation();

						// Hide.
							$pop4table._hide();

						// Redirect.
							window.setTimeout(function() {
								window.location.href = href;
							}, 350);

					});
//used to get to the menu by opening it once opened you can use escape to exit it, clicks to open and close
			$body
				.on('click', 'a[href="#pop4table"]', function(event) {

					event.stopPropagation();
					event.preventDefault();

					// Toggle.
						$pop4table._toggle();

				})
				.on('keydown', function(event) {

					// Hide on escape.
						if (event.keyCode == 27)
							$pop4table._hide();

				});

	});
$(function() {
		
			var $menu2 = $('#menu2');

			$menu2._locked = false;

			$menu2._lock = function() {

				if ($menu2._locked)
					return false;

				$menu2._locked = true;

				window.setTimeout(function() {
					$menu2._locked = false;
				}, 50);

				return true;

			};

			$menu2._show = function() {

				if ($menu2._lock())
					$body.addClass('is-menu2-visible');

			};

			$menu2._hide = function() {

				if ($menu2._lock())
					$body.removeClass('is-menu2-visible');

			};

			$menu2._toggle = function() {

				if ($menu2._lock())
					$body.toggleClass('is-menu2-visible');

			};

			$menu2
				$menu2.appendTo($body)
				.on('click', function(event) {

					event.stopPropagation();

					// Hide.
						$menu2._hide();

				})
				.find('.inner')
					.on('click', '.close', function(event) {

						event.preventDefault();
						event.stopPropagation();
						event.stopImmediatePropagation();

						// Hide.
							$menu2._hide();

					})
					.on('click', function(event) {
						event.stopPropagation();
					})
					.on('click', 'a', function(event) {

						var href = $(this).attr('href');

						event.preventDefault();
						event.stopPropagation();

						// Hide.
							$menu2._hide();

						// Redirect.
							window.setTimeout(function() {
								window.location.href = href;
							}, 350);

					});
//used to get to the menu by opening it once opened you can use escape to exit it, clicks to open and close
			$body
				.on('click', 'a[href="#menu2"]', function(event) {

					event.stopPropagation();
					event.preventDefault();

					// Toggle.
						$menu2._toggle();

				})
				.on('keydown', function(event) {

					// Hide on escape.
						if (event.keyCode == 27)
							$menu2._hide();

				});

	});
	//pop4form function
$(function() {
		
			var $pop4form = $('#pop4form');

			$pop4form._locked = false;

			$pop4form._lock = function() {

				if ($pop4form._locked)
					return false;

				$pop4form._locked = true;

				window.setTimeout(function() {
					$pop4form._locked = false;
				}, 50);

				return true;

			};

			$pop4form._show = function() {

				if ($pop4form._lock())
					$body.addClass('is-pop4form-visible');

			};

			$pop4form._hide = function() {

				if ($pop4form._lock())
					$body.removeClass('is-pop4form-visible');

			};

			$pop4form._toggle = function() {

				if ($pop4form._lock())
					$body.toggleClass('is-pop4form-visible');

			};

			$pop4form
				$pop4form.appendTo($body)
				.on('click', function(event) {

					event.stopPropagation();

					// Hide.
						$pop4form._hide();

				})
				.find('.inner')
					.on('click', '.close', function(event) {

						event.preventDefault();
						event.stopPropagation();
						event.stopImmediatePropagation();

						// Hide.
							$pop4form._hide();

					})
					.on('click', function(event) {
						event.stopPropagation();
					})
					.on('click', 'a', function(event) {

						var href = $(this).attr('href');

						event.preventDefault();
						event.stopPropagation();

						// Hide.
							$pop4form._hide();

						// Redirect.
							window.setTimeout(function() {
								window.location.href = href;
							}, 350);

					});
//used to get to the menu by opening it once opened you can use escape to exit it, clicks to open and close
			$body
				.on('click', 'a[href="#pop4form"]', function(event) {

					event.stopPropagation();
					event.preventDefault();

					// Toggle.
						$pop4form._toggle();

				})
				.on('keydown', function(event) {

					// Hide on escape.
						if (event.keyCode == 27)
							$pop4form._hide();

				});

	});


	$(function() {

		// Menu.
			var $menu = $('#menu');

			$menu._locked = false;

			$menu._lock = function() {

				if ($menu._locked)
					return false;

				$menu._locked = true;

				window.setTimeout(function() {
					$menu._locked = false;
				}, 50);

				return true;

			};

			$menu._show = function() {

				if ($menu._lock())
					$body.addClass('is-menu-visible');

			};

			$menu._hide = function() {

				if ($menu._lock())
					$body.removeClass('is-menu-visible');

			};

			$menu._toggle = function() {

				if ($menu._lock())
					$body.toggleClass('is-menu-visible');

			};

			$menu
				.appendTo($body)
				.on('click', function(event) {

					event.stopPropagation();

					// Hide.
						$menu._hide();

				})
				.find('.inner')
					.on('click', '.close', function(event) {

						event.preventDefault();
						event.stopPropagation();
						event.stopImmediatePropagation();

						// Hide.
							$menu._hide();

					})
					.on('click', function(event) {
						event.stopPropagation();
					})
					.on('click', 'a', function(event) {

						var href = $(this).attr('href');

						event.preventDefault();
						event.stopPropagation();

						// Hide.
							$menu._hide();

						// Redirect.
							window.setTimeout(function() {
								window.location.href = href;
							}, 350);

					});
//used to get to the menu by opening it once opened you can use escape to exit it, clicks to open and close
			$body
				.on('click', 'a[href="#menu"]', function(event) {

					event.stopPropagation();
					event.preventDefault();

					// Toggle.
						$menu._toggle();

				})
				.on('keydown', function(event) {

					// Hide on escape.
						if (event.keyCode == 27)
							$menu._hide();

				});
	});
	
	}catch(e){}})(jQuery);

//...
	function DropDown(el) {
    this.dd = el;
    this.placeholder = this.dd.children('span');
    this.opts = this.dd.find('ul.dropdown > li');
    this.val = '';
    this.index = -1;
    this.initEvents();
}