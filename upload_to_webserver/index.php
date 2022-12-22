<!DOCTYPE html>

<html lang="sv">

<head>
  <title>Statistik Elevfrånvaro</title>
  
  <meta charset="utf-8">
  
  <meta name="viewport" content="width=device-width, initial-scale=1">

  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-+0n0xVW2eSR5OomGNYDnhzAbDsOXxcvSN1TPprVMTNDbiYZCxYbOOl7+AMvyTG2x" crossorigin="anonymous">
  
  <link href="https://fonts.googleapis.com/css?family=Poppins:300,400,700,900&display=swap" rel="stylesheet">
  
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.3.1/css/all.css" integrity="sha384-mzrmE5qonljUremFsqc01SB46JvROS7bZs3IO2EmfFsd15uHvIt+Y8vEf7N7fWAU" crossorigin="anonymous">

  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
  
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/js/bootstrap.bundle.min.js" integrity="sha384-gtEjrD/SeCtmISkJkNUaaKMoLD0//ElJ19smozuHV6z3Iehds+3Ulb9Bn9Plx0x4" crossorigin="anonymous"></script>
  
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

  <link rel="shortcut icon" type="image/jpg" href="favicon.ico"/>

  <link id="theme" rel="stylesheet" type="text/css" href="style_color.css?ver=1.7" />

  <script>
  $(document).ready(function() {
      $(".divtime").load("time.php");
      var refreshId = setInterval(function() {
      $(".divtime").load("time.php");
      }, 60000);
      $(".divstatus").load("status.php");
      var refreshId2 = setInterval(function() {
      $(".divstatus").load("status.php");
      }, 60000);
      $(".sectionyear").load("total_and_sections.php");
      var refreshId3 = setInterval(function() {
      $(".sectionyear").load("total_and_sections.php");
      }, 5000000);
      $(".7weekhours").load("7weekhours.php");
      var refreshId4 = setInterval(function() {
      $(".7weekhours").load("7weekhours.php");
      }, 5000000);
      $(".79").load("79_today.php");
      var refreshId5 = setInterval(function() {
      $(".79").load("79_today.php");
      }, 60000);
      $(".f3").load("f3_today.php");
      var refreshId6 = setInterval(function() {
      $(".f3").load("f3_today.php");
      }, 69000);
      $(".46").load("46_today.php");
      var refreshId7 = setInterval(function() {
      $(".46").load("46_today.php");
      }, 60000);
      $(".allf9").load("total_today.php");
      var refreshId7 = setInterval(function() {
      $(".allf9").load("total_today.php");
      }, 60000);
      $(".noab").load("class_no_absence_today.php");
      var refreshId7 = setInterval(function() {
      $(".noab").load("class_no_absence_today.php");
      }, 60000);
      $(".hiab").load("class_high_absence_today.php");
      var refreshId7 = setInterval(function() {
      $(".hiab").load("class_high_absence_today.php");
      }, 60000);
      $(".baseinfo").load("base_info.php");
      var refreshId7 = setInterval(function() {
      $(".baseinfo").load("base_info.php");
      }, 60000);
      $(".res").load("resurs.php");
      var refreshId7 = setInterval(function() {
      $(".res").load("resurs.php");
      }, 60000);
      $(".allclass").load("all_classes.php");
      var refreshId7 = setInterval(function() {
      $(".allclass").load("all_classes.php");
      }, 60000);
      $(".15days").load("graph_15_days.php");
      var refreshId7 = setInterval(function() {
      $(".15days").load("graph_15_days.php");
      }, 60000);
  });
  </script>

</head>

<body>

<div class="container-fluid">
    
   <div class="row">
    
      <div class="col-md-3">
      
        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>STATISTIK </strong> ELEVFRÅNVARO</h2>
        <div class="divtime"></div>
        <div class="baseinfo"></div>
        <div class="res"></div> 
        </div></div>
        
        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>TOTALT</strong> LÄSÅRET 21/22</h2>
        <h4><strong>GENOMSNITTLIG</strong> FRÅNVARO</h4>
        <p class="infotextleft">Antal elever frånvarande per dag och klass - sorterat från högst till lägst - sedan läsårsstart</p>
        <div class="divstatus"></div></div></div>
     
      </div>
    
      <div class="col-md-9">

        <div class="row">

          <div class="col box4">
          <div class="allf9"></div>
          </div>

          <div class="col box5">
          <div class="noab"></div>
          </div>
          
          <div class="col box6">
          <div class="hiab"></div>
          </div>

          
        
        </div>

        <div class="row">

          <div class="col box1">
          <div class="f3"></div>
          </div>

          <div class="col box2">
          <div class="46"></div>
          </div>
          
          <div class="col box3">
          <div class="79"></div>
          </div>
        
        </div>
        
        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>SENASTE 15 DAGARNA</strong> TOTALT SAMT STADIER (%)</h2>
        <p class="infotext">Klicka på kategorierna för att visa eller gömma resultat</p>
        <div>
        <canvas id="myChart33"></canvas>
        <div class="15days"></div>
        </div>
        </div>
        </div>

        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>IDAG </strong> ANTAL FRÅNVARANDE ALLA KLASSER</h2>
        <div>
        <div class="allclass"></div>
        </div>
        </div>
        </div>

        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>LÄSÅRET 21/22</strong> TOTALT SAMT STADIER (%)</h2>
        <p class="infotext">Klicka på kategorierna för att visa eller gömma resultat</p>
        <div>
        <canvas id="myChart30"></canvas>
        <div class="sectionyear"></div>
        </div>
        </div>
        </div>
        
        <div class="row">
        <div class="div1-half-top">
        <h2 class="calheading"><strong>TRENDKURVA </strong> FÖR KLASSER MED HÖG FRÅNVARO IDAG</h2>
        <div>
        <canvas id="myChart1"></canvas>
        <p class="infotext">Funktion kommer i en senare uppdatering...</p>
        <div class="77weekhours"></div>
        </div>
        </div>
        </div>
   
      </div>

  </div>


</div>

</body>
</html>