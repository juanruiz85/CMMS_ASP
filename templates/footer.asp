<%
' =============================================================================
' CMMS - Template: Footer HTML
' Incluir al final de cada página protegida
' =============================================================================
%>
  </main>
</div><!-- /app-layout -->

<!-- JavaScript -->
<script src="/CMMS/assets/js/app.js"></script>

<!-- Scripts de página específica (definidos por cada módulo) -->
<% If IsDefined("PageScript") Then %>
<script>
<%= PageScript %>
</script>
<% End If %>

</body>
</html>
