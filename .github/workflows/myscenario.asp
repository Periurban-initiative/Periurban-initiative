	<LINK href="img/weblog2.css" type=text/css rel=stylesheet>
<% sub myscenario %>    
<DIV id=contentouter>
<DIV id=contentcontainer> 
  <DIV id=content> 
    <!-- Main Headline -->
    <!-- Side Bar 1 (Lower Left) -->
    <DIV class=floatcolmiddlehome> 
      <!-- Insert title here -->
      <BR>			<form action="editscenario.asp" method="post" name="data">

<TABLE cellSpacing=0 cellPadding=0 width=550 border=0>
        <TBODY>
          <TR height=24> 
            <TD vAlign=top height=24> <img src="img/center1.gif" border=0 height=24></TD>
          </TR>
          <TR> 
            <TD> 
			<TABLE cellSpacing=1 cellPadding=8 width=550 bgColor=#e47905 border=0>
                <TBODY>
                  <TR> 
                    <TD width=150 bgColor=white>  
                      <SPAN class=hed12px>Userid:</SPAN>
					</TD>
                    <TD width=400 bgColor=white>  
                      <SPAN class=hed12px>jeff</SPAN>
					</TD>
				  </TR>
                  <TR> 
                    <TD width=150 bgColor=white>  
                      <SPAN class=hed12px>Name:</SPAN>
					</TD>
                    <TD width=400 bgColor=white>  
                      <SPAN class=hed12px>Jeff Rosenblum</SPAN>
					</TD>
				  </TR>
                  <TR> 
                    <TD width=150 bgColor=white>  
                      <SPAN class=hed12px>Country:</SPAN>
					</TD>
                    <TD width=400 bgColor=white>  
                      <SPAN class=hed12px>
	<select name="title">
  <OPTION selected VALUE="<%=session("Title")%>"><%=session("Title")%>
  <OPTION VALUE="">None
  <OPTION VALUE="Afghanistan">Afghanistan
  <OPTION VALUE="Albania">Albania
  <OPTION VALUE="Algeria">Algeria
  <OPTION VALUE="American Samoa">American Samoa
  <OPTION VALUE="Andorra">Andorra
  <OPTION VALUE="Angola">Angola
  <OPTION VALUE="Anguilla">Anguilla
  <OPTION VALUE="Antarctica">Antarctica
  <OPTION VALUE="Antigua and Barbuda">Antigua and Barbuda
  <OPTION VALUE="Argentina">Argentina
  <OPTION VALUE="Armenia">Armenia
  <OPTION VALUE="Aruba">Aruba
  <OPTION VALUE="Australia">Australia
  <OPTION VALUE="Austria">Austria
  <OPTION VALUE="Azerbaijan">Azerbaijan
  <OPTION VALUE="Bahamas">Bahamas
  <OPTION VALUE="Bahrain">Bahrain
  <OPTION VALUE="Bangladesh">Bangladesh
  <OPTION VALUE="Barbados">Barbados
  <OPTION VALUE="Belarus">Belarus
  <OPTION VALUE="Belgium">Belgium
  <OPTION VALUE="Belize">Belize
  <OPTION VALUE="Benin">Benin
  <OPTION VALUE="Bermuda">Bermuda
  <OPTION VALUE="Bhutan">Bhutan
  <OPTION VALUE="Bolivia">Bolivia
  <OPTION VALUE="Bosnia and Herzegovina">Bosnia and 
              Herzegovina
  <OPTION VALUE="Botswana">Botswana
  <OPTION VALUE="Bouvet Island">Bouvet Island
  <OPTION VALUE="Brazil">Brazil
  <OPTION VALUE="British Indian Ocean Territory">
              British Indian Ocean Territory
  <OPTION VALUE="Brunei Darussalam">Brunei Darussalam
  <OPTION VALUE="Bulgaria">Bulgaria
  <OPTION VALUE="Burkina Faso">Burkina Faso
  <OPTION VALUE="Burundi">Burundi
  <OPTION VALUE="Cambodia">Cambodia
  <OPTION VALUE="Cameroon">Cameroon
  <OPTION VALUE="Canada">Canada
  <OPTION VALUE="Cape Verde">Cape Verde
  <OPTION VALUE="Cayman Islands">Cayman Islands
  <OPTION VALUE="Central African Republic">
             Central African Republic
  <OPTION VALUE="Chad">Chad
  <OPTION VALUE="Chile">Chile
  <OPTION VALUE="China">China
  <OPTION VALUE="Christmas Island">Christmas Island
  <OPTION VALUE="Cocos (Keeling Islands)">
             Cocos (Keeling Islands)
  <OPTION VALUE="Colombia">Colombia
  <OPTION VALUE="Comoros">Comoros
  <OPTION VALUE="Congo">Congo
  <OPTION VALUE="Cook Islands">Cook Islands
  <OPTION VALUE="Costa Rica">Costa Rica
  <OPTION VALUE="Cote D'Ivoire (Ivory Coast)">
               Cote D'Ivoire (Ivory Coast)
  <OPTION VALUE="Croatia (Hrvatska">Croatia (Hrvatska
  <OPTION VALUE="Cuba">Cuba
  <OPTION VALUE="Cyprus">Cyprus
  <OPTION VALUE="Czech Republic">Czech Republic
  <OPTION VALUE="Denmark">Denmark
  <OPTION VALUE="Djibouti">Djibouti
  <OPTION VALUE="Dominican Republic">Dominican Republic
  <OPTION VALUE="Dominica">Dominica
  <OPTION VALUE="East Timor">East Timor
  <OPTION VALUE="Ecuador">Ecuador
  <OPTION VALUE="Egypt">Egypt
  <OPTION VALUE="El Salvador">El Salvador
  <OPTION VALUE="Equatorial Guinea">Equatorial Guinea
  <OPTION VALUE="Eritrea">Eritrea
  <OPTION VALUE="Estonia">Estonia
  <OPTION VALUE="Ethiopia">Ethiopia
  <OPTION VALUE="Falkland Islands (Malvinas)">
                  Falkland Islands (Malvinas)
  <OPTION VALUE="Faroe Islands">Faroe Islands
  <OPTION VALUE="Fiji">Fiji
  <OPTION VALUE="Finland">Finland
  <OPTION VALUE="France, Metropolitan">France, Metropolitan
  <OPTION VALUE="France">France
  <OPTION VALUE="French Guiana">French Guiana
  <OPTION VALUE="French Polynesia">French Polynesia
  <OPTION VALUE="French Southern Territories">
              French Southern Territories
  <OPTION VALUE="Gabon">Gabon
  <OPTION VALUE="Gambia">Gambia
  <OPTION VALUE="Georgia">Georgia
  <OPTION VALUE="Germany">Germany
  <OPTION VALUE="Ghana">Ghana
  <OPTION VALUE="Gibraltar">Gibraltar
  <OPTION VALUE="Greece">Greece
  <OPTION VALUE="Greenland">Greenland
  <OPTION VALUE="Grenada">Grenada
  <OPTION VALUE="Guadeloupe">Guadeloupe
  <OPTION VALUE="Guam">Guam
  <OPTION VALUE="Guatemala">Guatemala
  <OPTION VALUE="Guinea-Bissau">Guinea-Bissau
  <OPTION VALUE="Guinea">Guinea
  <OPTION VALUE="Guyana">Guyana
  <OPTION VALUE="Haiti">Haiti
  <OPTION VALUE="Heard and McDonald Islands">
            Heard and McDonald Islands
  <OPTION VALUE="Honduras">Honduras
  <OPTION VALUE="Hong Kong">Hong Kong
  <OPTION VALUE="Hungary">Hungary
  <OPTION VALUE="Iceland">Iceland
  <OPTION VALUE="India">India
  <OPTION VALUE="Indonesia">Indonesia
  <OPTION VALUE="Iran">Iran
  <OPTION VALUE="Iraq">Iraq
  <OPTION VALUE="Ireland">Ireland
  <OPTION VALUE="Israel">Israel
  <OPTION VALUE="Italy">Italy
  <OPTION VALUE="Jamaica">Jamaica
  <OPTION VALUE="Japan">Japan
  <OPTION VALUE="Jordan">Jordan
  <OPTION VALUE="Kazakhstan">Kazakhstan
  <OPTION VALUE="Kenya">Kenya
  <OPTION VALUE="Kiribati">Kiribati
  <OPTION VALUE="Korea (North)">Korea (North)
  <OPTION VALUE="Korea (South)">Korea (South)
  <OPTION VALUE="Kuwait">Kuwait
  <OPTION VALUE="Kyrgyzstan">Kyrgyzstan
  <OPTION VALUE="Laos">Laos
  <OPTION VALUE="Latvia">Latvia
  <OPTION VALUE="Lebanon">Lebanon
  <OPTION VALUE="Lesotho">Lesotho
  <OPTION VALUE="Liberia">Liberia
  <OPTION VALUE="Libya">Libya
  <OPTION VALUE="Liechtenstein">Liechtenstein
  <OPTION VALUE="Lithuania">Lithuania
  <OPTION VALUE="Luxembourg">Luxembourg
  <OPTION VALUE="Macau">Macau
  <OPTION VALUE="Macedonia">Macedonia
  <OPTION VALUE="Madagascar">Madagascar
  <OPTION VALUE="Malawi">Malawi
  <OPTION VALUE="Malaysia">Malaysia
  <OPTION VALUE="Maldives">Maldives
  <OPTION VALUE="Mali">Mali
  <OPTION VALUE="Malta">Malta
  <OPTION VALUE="Marshall Islands">Marshall Islands
  <OPTION VALUE="Martinique">Martinique
  <OPTION VALUE="Mauritania">Mauritania
  <OPTION VALUE="Mauritius">Mauritius
  <OPTION VALUE="Mayotte">Mayotte
  <OPTION VALUE="Mexico">Mexico
  <OPTION VALUE="Micronesia">Micronesia
  <OPTION VALUE="Moldova">Moldova
  <OPTION VALUE="Monaco">Monaco
  <OPTION VALUE="Mongolia">Mongolia
  <OPTION VALUE="Montserrat">Montserrat
  <OPTION VALUE="Morocco">Morocco
  <OPTION VALUE="Mozambique">Mozambique
  <OPTION VALUE="Myanmar">Myanmar
  <OPTION VALUE="Namibia">Namibia
  <OPTION VALUE="Nauru">Nauru
  <OPTION VALUE="Nepal">Nepal
  <OPTION VALUE="Netherlands Antilles">Netherlands Antilles
  <OPTION VALUE="Netherlands">Netherlands
  <OPTION VALUE="New Caledonia">New Caledonia
  <OPTION VALUE="New Zealand">New Zealand
  <OPTION VALUE="Nicaragua">Nicaragua
  <OPTION VALUE="Nigeria">Nigeria
  <OPTION VALUE="Niger">Niger
  <OPTION VALUE="Niue">Niue
  <OPTION VALUE="Norfolk Island">Norfolk Island
  <OPTION VALUE="Northern Mariana Islands">
             Northern Mariana Islands
  <OPTION VALUE="Norway">Norway
  <OPTION VALUE="Oman">Oman
  <OPTION VALUE="Pakistan">Pakistan
  <OPTION VALUE="Palau">Palau
  <OPTION VALUE="Panama">Panama
  <OPTION VALUE="Papua New Guinea">Papua New Guinea
  <OPTION VALUE="Paraguay">Paraguay
  <OPTION VALUE="Peru">Peru
  <OPTION VALUE="Philippines">Philippines
  <OPTION VALUE="Pitcairn">Pitcairn
  <OPTION VALUE="Poland">Poland
  <OPTION VALUE="Portugal">Portugal
  <OPTION VALUE="Puerto Rico">Puerto Rico
  <OPTION VALUE="Qatar">Qatar
  <OPTION VALUE="Reunion">Reunion
  <OPTION VALUE="Romania">Romania
  <OPTION VALUE="Russian Federation">Russian Federation
  <OPTION VALUE="Rwanda">Rwanda
  <OPTION VALUE="S. Georgia and S. Sandwich Isls.">
         S. Georgia and S. Sandwich Isls.
  <OPTION VALUE="Saint Kitts and Nevis">Saint Kitts and Nevis
  <OPTION VALUE="Saint Lucia">Saint Lucia
  <OPTION VALUE="Saint Vincent and The Grenadines">
         Saint Vincent and The Grenadines
  <OPTION VALUE="Samoa">Samoa
  <OPTION VALUE="San Marino">San Marino
  <OPTION VALUE="Sao Tome and Principe">Sao Tome and Principe
  <OPTION VALUE="Saudi Arabia">Saudi Arabia
  <OPTION VALUE="Senegal">Senegal
  <OPTION VALUE="Seychelles">Seychelles
  <OPTION VALUE="Sierra Leone">Sierra Leone
  <OPTION VALUE="Singapore">Singapore
  <OPTION VALUE="Slovak Republic">Slovak Republic
  <OPTION VALUE="Slovenia">Slovenia
  <OPTION VALUE="Solomon Islands">Solomon Islands
  <OPTION VALUE="Somalia">Somalia
  <OPTION VALUE="South Africa">South Africa
  <OPTION VALUE="Spain">Spain
  <OPTION VALUE="Sri Lanka">Sri Lanka
  <OPTION VALUE="St. Helena">St. Helena
  <OPTION VALUE="St. Pierre and Miquelon">
              St. Pierre and Miquelon
  <OPTION VALUE="Sudan">Sudan
  <OPTION VALUE="Suriname">Suriname
  <OPTION VALUE="Svalbard and Jan Mayen Islands">
              Svalbard and Jan Mayen Islands
  <OPTION VALUE="Swaziland">Swaziland
  <OPTION VALUE="Sweden">Sweden
  <OPTION VALUE="Switzerland">Switzerland
  <OPTION VALUE="Syria">Syria
  <OPTION VALUE="Taiwan">Taiwan
  <OPTION VALUE="Tajikistan">Tajikistan
  <OPTION VALUE="Tanzania">Tanzania
  <OPTION VALUE="Thailand">Thailand
  <OPTION VALUE="Togo">Togo
  <OPTION VALUE="Tokelau">Tokelau
  <OPTION VALUE="Tonga">Tonga
  <OPTION VALUE="Trinidad and Tobago">Trinidad and Tobago
  <OPTION VALUE="Tunisia">Tunisia
  <OPTION VALUE="Turkey">Turkey
  <OPTION VALUE="Turkmenistan">Turkmenistan
  <OPTION VALUE="Turks and Caicos Islands">
       Turks and Caicos Islands
  <OPTION VALUE="Tuvalu">Tuvalu
  <OPTION VALUE="US Minor Outlying Islands">
     US Minor Outlying Islands
  <OPTION VALUE="Uganda">Uganda
  <OPTION VALUE="Ukraine">Ukraine
  <OPTION VALUE="United Arab Emirates">
     United Arab Emirates
  <OPTION VALUE="United Kingdom">United Kingdom
  <OPTION VALUE="United States">United States
  <OPTION VALUE="Uruguay">Uruguay
  <OPTION VALUE="Uzbekistan">Uzbekistan
  <OPTION VALUE="Vanuatu">Vanuatu
  <OPTION VALUE="Vatican City State">Vatican City State
  <OPTION VALUE="Venezuela">Venezuela
  <OPTION VALUE="Viet Nam">Viet Nam
  <OPTION VALUE="Virgin Islands (British)">
     Virgin Islands (British)
  <OPTION VALUE="Virgin Islands (US)">
     Virgin Islands (US)
  <OPTION VALUE="Wallis and Futuna Islands">
     Wallis and Futuna Islands
  <OPTION VALUE="Western Sahara">Western Sahara
  <OPTION VALUE="Yemen">Yemen
  <OPTION VALUE="Yugoslavia">Yugoslavia
  <OPTION VALUE="Zaire">Zaire
  <OPTION VALUE="Zambia">Zambia
  <OPTION VALUE="Zimbabwe">Zimbabwe
</SELECT>					  
					  </SPAN>
					</TD>
				  </TR>
                  <TR> 
                    <TD width=150 bgColor=white>  
                      <SPAN class=hed12px>Notes:</SPAN>
					</TD>
                    <TD width=400 bgColor=white>  
                      <SPAN class=hed12px><TEXTAREA NAME="desc" ROWS=10 COLS=50 STYLE="font-size : 10pt" value=""><%=session("Desc")%></textarea></SPAN>
					</TD>
				  </TR>
					
                </TBODY>
              </TABLE></TD>
          </TR></TD></TR>
        </TBODY>
      </TABLE>
	  <BR><INPUT type="submit" value="SUBMIT CHANGES">
	  </form>
    </DIV>
  </DIV>
</DIV> </div>
<% end sub %>