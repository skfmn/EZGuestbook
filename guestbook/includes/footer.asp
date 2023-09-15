<!-- Footer -->
	<footer id="footer">
		<div class="copyright">
			<a href="http://www.aspjunction.com">EZGuestbook</a> Copyright &copy; 2003 - <%= Year(Date) %> <a href="http://www.aspjunction.com">ASP junction</a>
		</div>
	</footer>

  <style>
    div.ui-menu-item-wrapper {color:#ff0000;}
    div.ui-helper-hidden-accessible:last-child {color:#ff0000;}
  </style>

  <!-- Scripts -->
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  <script src="../assets/js/jquery.fancybox.js"></script>
  <script src="../assets/js/skel.min.js"></script>
  <script src="../assets/js/main.js"></script>
  <script src="../assets/js/javascript.js"></script>
  <script language="javascript" type="text/javascript">
  $(document).ready(function(){
	  $(".iframe").fancybox();
	  $(".picimg").fancybox();
	  $("#textmsg").fancybox({border:0});
	  $("#textmsg").trigger('click');
  }); 
  </script> 
</body>
</html>
