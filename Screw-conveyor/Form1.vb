Imports System.Math
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms

Public Structure Conveyor_struct    'For conveyors
    Public Tag As String            '[T4400]
    Public no_screws As Integer     '[-]
    Public flight_OD As Double     '[mm] 
    Public pipe_od As Double        '[mm]
    Public pipe_ID As Double        '[mm]
    Public pitch As Double          '[mm]
    Public flight_thick As Double   '[mm]
    Public flight_weight As Double  '[kg] weight of one 360 degree flight
    Public cap_1rev As Double       '[ltr/rev]
    Public cap_sys As Double        '[kg/hr] capacity system
    Public density As Double        '[kg/m3]
    Public filling As Double        '[%]
    Public rpm As Double            '[rpm]
End Structure

Public Class Form1
    'Use icon convert site https://icoconvert.com/ 
    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Conveyor_sizing_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Conveyor_rapport_copy\"
    Dim dirpath_Home_GP As String = "C:\Temp\"

    Public conv As Conveyor_struct   'Conveyors data

    Public _steps As Integer = 150   'Calculation _steps
    Public _d(_steps) As Double      '[m] Distance to drive plate
    Public _s(_steps) As Double      '[N] Shear force @ distance to drive plate
    Public _m(_steps) As Double      '[Nm] Moment  @ distance to drive plate
    Public _α(_steps) As Double      '[rad] Deflection angle @ distance to drive plate
    Public _αv(_steps) As Double     '[rad] Deflection @ distance to drive plate

    '-------- inlet chute dimensions ---
    Public _κ1 As Double     '[m] exposed pipe length force 1
    Public _κ2 As Double     '[m] exposed pipe length force 2
    Public _κ3 As Double     '[m] exposed pipe length force 3

    '-------- conveyor dimensions--------
    Public _λ1 As Double     '[m] drive shaft length   
    Public _λ2 As Double     '[m] center inlet chute #1   
    Public _λ3 As Double     '[m] center inlet chute #2   
    Public _λ4 As Double     '[m] center inlet chute #3     
    Public _λ5 As Double     '[m] tail shaft length   
    Public _λ6 As Double     '[m] flight/trough length
    Public _λ7 As Double     '[m] bearing-bearing length

    Public Shared _diam_flight As Double                         '[m]
    Public Shared _pipe_OD, _pipe_ID, _pipe_wall As Double
    Public Shared pipe_Ix, pipe_Wx, pipe_Wp As Double            'Lineair en polair weerstand moment
    Public Shared pitch As Double
    Public Shared installed_power As Double
    Public Shared Start_factor As Double
    Public Shared actual_power As Double
    Public Shared sigma02, sigma_fatique As Double
    Public Young As Double = 210000

    Public Shared inlet_length, product_density As Double

    Public Shared _angle As Double
    Public Shared speed As Double
    Public Shared _regu_flow_kg_hr As Double
    Public Shared density As Double

    Public Shared progress_resistance As Double    'Friction from product to steel

    'Materials name; CEMA Material code; Conveyor loading; Component group, density min, Density max, HP Material
    Public Shared _inputs() As String = {
" 300 [kg/m3]; 0000;30A;2B;300;300;1.0",
" 350 [kg/m3]; 0000;30A;2B;350;350;1.0",
" 400 [kg/m3]; 0000;30A;2B;400;400;1.0",
" 450 [kg/m3]; 0000;30A;2B;450;450;1.0",
" 500 [kg/m3]; 0000;30A;2B;500;500;1.0",
" 550 [kg/m3]; 0000;30A;2B;550;550;1.0",
" 600 [kg/m3]; 0000;30A;2B;600;600;1.0",
" 650 [kg/m3]; 0000;30A;2B;650;650;1.0",
" 700 [kg/m3]; 0000;30A;2B;700;700;1.0",
" 750 [kg/m3]; 0000;30A;2B;750;750;1.0",
" 800 [kg/m3]; 0000;30A;2B;800;800;1.0",
" 850 [kg/m3]; 0000;30A;2B;850;850;1.0",
" 900 [kg/m3]; 0000;30A;2B;900;900;1.0",
" 950 [kg/m3]; 0000;30A;2B;950;950;1.0",
"1000 [kg/m3]; 0000;30A;2B;1000;1000;1.0",
"1100 [kg/m3]; 0000;30A;2B;1100;1100;1.0",
"1200 [kg/m3]; 0000;30A;2B;1200;1200;1.0",
"1300 [kg/m3]; 0000;30A;2B;1300;1300;1.0",
"1400 [kg/m3]; 0000;30A;2B;1400;1400;1.0",
"1500 [kg/m3]; 0000;30A;2B;1500;1500;1.0",
"1600 [kg/m3]; 0000;30A;2B;1600;1600;1.0",
"1700 [kg/m3]; 0000;30A;2B;1700;1700;1.0",
"1800 [kg/m3]; 0000;30A;2B;1800;1800;1.0",
"1900 [kg/m3]; 0000;30A;2B;1900;1900;1.0",
"2000 [kg/m3]; 0000;30A;2B;2000;2000;1.0",
"Adipic-Acid;45A35;30A;2B;720;720;0.5",
"Alfalfa Meal;18B45WY;30A;2D;220;350;0.6",
"Alfalfa Pellets;42C25;45;2D;660;690;0.5",
"Alfalfa Seed;13B15N;45;1A,1B,1C;160;240;0.4",
"Almonds Broken;29C35Q;30A;2D;430;480;0.9",
"Almonds Whole Shelled;29C35Q;30A;2D;450;480;0.9",
"Alum Fine;48B35U;30A;3D;720;800;0.6",
"Alum, Lumps;55B25;45;2A,2B;800;960;1.4",
"Alumina;58B27MY;15;3D;880;1040;1.8",
"Alumina Fines;35A27MY;15;3D;560;560;1.6",
"Alumina Sized or Briquette;65D37;15;3D;1040;1040;2",
"Aluminate Gel (Aluminate Hydroxide);45B35;30B;2D;720;720;1.7",
"Aluminum Chips, Dry;11E45V;30A;2D;110;240;1.2",
"Aluminum Chips, Oily;11E45VY;30A;2D;110;240;0.8",
"Aluminum Hydrate;17C35;30A;1A,1B,1C;210;320;1.4",
"Aluminum Oxide;90A17MN;15;3D;960;1920;1.8",
"Aluminum Silicate (Andalusite);49C35S;45;3A,3B;780;780;0.7",
"Aluminum Sulfate;52C25;45;1A,1B,1C;720;930;1.3",
"Ammonium Chloride, Crystalline;49A45FRS;30A;1A,1B,1C;720;830;1",
"Ammonium Nitrate;54A35NTU;30A;3D;720;990;1.6",
"Ammonium Sulfate;52C35FOTU;30A;1A,1B,1C;720;930;1",
"Apple Pomace, Dry;15C45Y;30B;2D;240;240;1",
"Arsenate of Lead (Lead Arsenate);72A35R;30A;1A,1B,1C;1150;1150;1.4",
"Arsenic Pulverized;30A25R;45;2D;480;480;1",
"Asbestos-Rock (Ore);81D37R;15;3D;1300;1300;2",
"Asbestos-Shredded;30E46XY;30B;2D;320;640;1",
"Ash, Black Ground;105B35;30A;1A,1B,1C;1680;1680;2.5",
"Ashes, Coal, dry, 1/2 inch;40C46TY;30B;3D;560;720;3",
"Ashes, Coal, dry, 3 inch;38D46T;30B;3D;560;640;2.5",
"Ashes, Coal, Wet, 1/2 inch;48C46T;30B;3D;720;800;3",
"Ashes, Coal, Wet, 3 inch;48D46T;30B;3D;720;800;4",
"Ashes, Fly (Fly Ash);38A36M;30B;3D;480;720;2",
"Aspartic Acid;42A35XPLO;30A;1A,1B,1C;530;820;1.5",
"Asphalt, Crushed, 1/2 inch;45C45;30A;1A,1B,1C;720;720;2",
"Bagasse;9E45RVXY;30A;2A,2B,2C;110;160;1.5",
"Bakelite, Fine;38B25;45;1A,1B,1C;480;720;1.4",
"Baking Powder;48A35;30A;1B;640;880;0.6",
"Baking Soda (Sodium Bicarbonate);48A25;45;1B;640;880;0.6",
"Barite (Barium Sulfate), 1/2 to 3 inch;150D36;30B;3D;1920;2880;2.6",
"Barite, Powder;150A35X;30A;2D;1920;2880;2",
"Barium Carbonate;72A45R;30A;2D;1150;1150;1.6",
"Bark Wood, Refuse;15E45TVY;30A;3D;160;320;2",
"Barley Fine, Ground;31B35;30A;1A,1B,1C;380;610;0.4",
"Barley Malted;31C35;30A;1A,1B,1C;500;500;0.4",
"Barley Meal;28C35;30A;1A,1B,1C;450;450;0.4",
"Barley Whole;42B25N;45;1A,1B,1C;580;770;0.5",
"Basalt;93B27;15;3D;1280;1680;1.8",
"Bauxite, Crushed, 3 inch (Aluminum Ore);80D36;30B;3D;1200;1360;2.5",
"Bauxite Dry, Ground(Aluminum Ore);68B25;45;2D;1090;1090;1.8",
"Beans Castor, Meal;38B35W;30A;1A,1B,1C;560;640;0.8",
"Beans Castor, Whole Shelled;36C15W;45;1A,1B,1C;580;580;0.5",
"Beans Navy, Dry;48C15;45;1A,1B,1C;770;770;0.5",
"Beans Navy, Steeped;60C25;45;1A,1B,1C;960;960;0.8",
"Bentonite 100 Mesh;55A25MXY;45;2D;800;960;0.7",
"Bentonite Crude;37D45X;30A;2D;540;640;1.2",
"Benzene Hexachloride;56A45R;30A;1A,1B,1C;900;900;0.6",
"Bicarbonate of Soda (Baking Soda);48A25;45;1B;640;880;0.6",
"Blood Dried;40D45U;30A;2D;560;720;2",
"Blood Ground, Dried;30A35U;30A;1A,1B;480;480;1",
"Bone Ash (Tricalcium Phosphate);45A45;30A;1A,1B;640;800;1.6",
"Boneblack;23A25Y;45;1A,1B;320;400;1.5",
"Bonechar;34B35;30A;1A,1B;430;640;1.6",
"Bonemeal;55B35;30A;2D;800;960;1.7",
"Bones Crushed;43D45;30A;2D;560;800;2",
"Bones Ground;50B35;30A;2D;800;800;1.7",
"Bones Whole**;43E45V;30A;2D;560;800;3",
"Borate of Lime;60A35;30A;1A,1B,1C;960;960;0.6",
"Borax Screening, 1/2 inch;58C35;30A;2D;880;960;1.5",
"Borax 1-1/2  to 2 inch Lump;58D35;30A;2D;880;960;1.8",
"Borax 2 to 3 inch Lump;65D35;30A;2D;960;1120;2",
"Borax Fine;50B25T;45;3D;720;880;0.7",
"Boric Acid, Fine;55B25T;45;3D;880;880;0.8",
"Boron;75A37;15;2D;1200;1200;1",
"Bran, Rice-Rye-Wheat;18B355NY;30A;1A,1B,1C;260;320;0.5",
"Braunite (Manganese Oxide);120A36;30B;2D;1920;1920;2",
"Bread Crumbs;23B35PQ;30A;1A,1B,1C;320;400;0.6",
"Brewers Grain, spent, dry;22C45;30A;1A,1B,1C;220;480;0.5",
"Brewers Grain, spent, wet;58C45T;30A;2A,2B;880;960;0.8",
"Brick, Ground, 1/8 inch;110B37;15;3D;1600;1920;2.2",
"Bronze Chips;40B45;30A;2D;480;800;2",
"Buckwheat;40B25N;45;1A,1B,1C;590;670;0.4",
"Calcine, Flour;80A35;30A;1A,1B,1C;1200;1360;0.7",
"Calcium Carbide;80D25N;30A;2D;1120;1440;2",
"Calcium Hydrate (Lime, Hydrated);40B35LM;30A;2D;640;640;0.8",
"Calcium Hydroxide (Lime, Hydrated);40B35LM;30A;2D;640;640;0.8",
"Calcium Lactate;28D45QTR;30A;2A,2B;420;460;0.6",
"Calcium Oxide (Lime, unslaked);63B35U;30A;1A,1B,1C;960;1040;0.6",
"Calcium Phosphate;45A45;30A;1A,1B,1C;640;800;1.6",
"Canola Meal (Rape Seed Meal)**;38;?;?;540;660;0.8",
"Carborundum;100D27;15;3D;1600;1600;3",
"Casein;36B35;30A;2D;580;580;1.6",
"Cashew Nuts;35C45;30A;2D;510;590;0.7",
"Cast Iron, Chips;165C45;30A;2D;2080;3200;4",
"Caustic Soda (Sodium Hydroxide);88B35RSU;30A;3D;1410;1410;1.8",
"Caustic Soda, Flakes;47C45RSUX;30A;3A,3B;750;750;1.5",
"Celite (Diatomaceous Earth);14A36Y;30B;3D;180;270;1.6",
"Cellulose with TBA;VTK;30B;2D;960;800;1.6",
"Cement, Aerated (Portland);68A16M;30B;2D;960;1200;1.4",
"Cement, Clinker;85D36;30B;3D;1200;1520;1.8",
"Cement, Mortar;133B35Q;30A;3D;2130;2130;3",
"Cement, Portland;94A26M;30B;2D;1510;1510;1.4",
"Cerrusite (Lead Carbonate);250A35R;30A;2D;3840;4160;1",
"Chalk, Crushed;85D25;30A;2D;1200;1520;1.9",
"Chalk, Pulverized;71A25MXY;45;2D;1070;1200;1.4",
"Charcoal, Ground;23A45;30A;2D;290;450;1.2",
"Charcoal, Lumps;23D45Q;30A;2D;290;450;1.4",
"Chocolate, Cake Pressed;43D25;30A;2B;640;720;1.5",
"Chrome Ore;133D36;30B;3D;2000;2240;2.5",
"Cinders, Blast Furnace;57D36T;30B;3D;910;910;1.9",
"Cinders, Coal;40D36T;30B;3D;640;640;1.8",
"Clay (Marl);80D36;30B;2D;1280;1280;1.6",
"Clay, Brick, Dry, Fines;110C36;30B;3D;1600;1920;2",
"Clay, Calcined;90B36;30B;3D;1280;1600;2.4",
"Clay, Ceramic, Dry, Fines;70A35P;30A;1A,1B,1C;960;1280;1.5",
"Clay, Dry, Lumpy;68D35;30A;2D;960;1200;1.8",
"Clinker, Cement (Cement Clinker);85D36;30B;3D;1200;1520;1.8",
"Clover Seed;47B25N;45;1A,1B,1C;720;770;0.4",
"Coal, Anthracite (River & Culm);58B35TY;30A;2A,2B;880;980;1",
"Coal, Anthracite, Sized, 1/2 inch;55C25;45;2A,2B;780;980;1",
"Coal, Bituminous, Mined;50D35LNYX;30A;1A,1B;640;960;1",
"Coal, Bituminous, Mined, Sized;48D35QV;30A;1A,1B;720;800;1",
"Coal, Bituminous, Mined, Slack;47C45T;30A;2A,2B;690;800;0.9",
"Coal, Lignite;41D35T;30A;2D;590;720;1",
"Cocoa Beans;38C25Q;30A;1A,1B;480;720;0.5",
"Cocoa, Nibs;35C25;45;2D;560;560;0.5",
"Cocoa, Powdered;33A45XY;30A;1B;480;560;0.9",
"Coconut, Shredded;2.1E+46;30B;2B;320;350;1.5",
"Coffee, Chaff;20B25FZMY;45;1A,1B;320;320;1",
"Coffee, Green Bean;29C25PQ;45;1A,1B;400;510;0.5",
"Coffee, Ground, Dry;25A35P;30A;1A,1B;400;400;0.6",
"Coffee, Ground, Wet;40A45X;30A;1A,1B;560;720;0.6",
"Coffee, Roasted Bean;25C25PQ;45;1B;320;480;0.4",
"Coffee, Soluble;19A35PUY;30A;1B;300;300;0.4",
"Coke, Breeze;30C37;15;3D;400;560;1.2",
"Coke, Loose;30D37;15;3D;400;560;1.2",
"Coke, Petrol, Calcined;40D37;15;3D;560;720;1.3",
"Compost;40D45TV;30A;3A,3B;480;800;1",
"Concrete, Pre-Mix,;103C36U;30B;3D;1360;1920;3",
"Copper Ore;135D36;30B;3D;1920;2400;4",
"Copper Ore, Crushed;125D36;30B;3D;1600;2400;4",
"Copper Sulphate, (Bluestone, Cupric Sulphate);85C35S;30A;2A,2B,2C;1200;1520;1",
"Copperas (Ferrous Sulphate);63C35U;30A;2D;800;1200;1",
"Copra, Cake Ground;43B45HW;30A;1A,1B,1C;640;720;0.7",
"Copra, Cake, Lumpy;28D35HW;30A;2A,2B,2C;400;480;0.8",
"Copra, Lumpy;22E35HW;30A;2A,2B,2C;350;350;1",
"Copra, Meal;43B35HW;30A;2D;640;720;0.7",
"Cork, Fine Ground;10B35JNY;30A;1A,1B,1C;80;240;0.5",
"Cork, Granulated;14C35JY;30A;1A,1B,1C;190;240;0.5",
"Corn Cobs, Ground;17C25Y;45;1A,1B,1C;270;270;0.6",
"Corn Fiber, Dry;14B46P;30B;1A,1B,1C;190;240;1",
"Corn Fiber, Wet;33B46P;30B;1A,1B,1C;240;800;1.5",
"Corn Oil, Cake;25D45HW;30A;1A,1B;400;400;0.6",
"Corn, Cracked;45B25P;45;1A,1B,1C;640;800;0.7",
"Corn, Germ, Dry;21B35PY;30A;1A,1B,1C;340;340;0.4",
"Corn, Germ, Wet (50%, moisture);30B35PY;30A;1A,1B,1C;480;480;0.4",
"Corn, Grits;43B35P;30A;1A,1B,1C;640;720;0.5",
"Corn, Seed;45C25PQ;45;1A,1B,1C;720;720;0.4",
"Corn, Shelled;45C25;45;1A,1B,1C;720;720;0.4",
"Corn, Starch*;38A15MN;45;1A,1B,1C;400;800;1",
"Corn, Sugar;33B35PU;30A;1B;480;560;1.3",
"Cornmeal;36B35P;30A;1A,1B;510;640;0.5",
"Cottonseed, Cake;43C45HW;30A;1A,1B;640;720;1",
"Cottonseed, Dry, Delinted;31C25X;45;1A,1B;350;640;0.6",
"Cottonseed, Dry, Not Delinted;22C45XY;30A;1A,1B;290;400;0.9",
"Cottonseed, Flakes;23C35HWY;30A;1A,1B;320;400;0.8",
"Cottonseed, Hulls;12B35Y;30A;1A,1B;190;190;0.9",
"Cottonseed, Meal, Expeller;28B45HW;30A;3A,3B;400;480;0.5",
"Cottonseed, Meal, Extracted;38B45HW;30A;1A,1B;560;640;0.5",
"Cottonseed, Meats, Dry;40B35HW;30A;1A,1B;640;640;0.6",
"Cottonseed, Meats, Rolled;38C45HW;30A;1A,1B;560;640;0.6",
"Cracklings, Crushed;45D45HW;30A;2A,2B,2C;640;800;1.3",
"Cryolite, Dust (Sodium Aluminum Fluoride);83A36V;30B;2D;1200;1440;2",
"Cryolite, Lumpy (Kryalith);100D36;30B;2D;1440;1760;2.1",
"Cullet, Fine;100C37;15;3D;1280;1920;2",
"Cullet, Lump;100D37;15;3D;1280;1920;2.5",
"Culm, (Coal, Anthracite);58B35TY;30A;2A,2B;880;980;1",
"Cupric Sulphate (Copper Sulfate);85C35S;30A;2A,2B,2C;1200;1520;1",
"Diatomaceous Earth (Filter Aid, Precoat);14A36Y;30B;3D;180;270;1.6",
"Dicalcium Phosphate;45A35;30A;1A,1B,1C;640;800;1.6",
"Disodium Phosphate;28A35;30A;3D;400;500;0.5",
"Distillers Grain, Spent Wet;50C45V;30A;3A,3B;640;960;0.8",
"Distillers Grain, Spent Wet w/Syrup;56C45VXOH;30A;3A,3B;690;1090;1.2",
"Distillers Grain-Spent Dry;30B35;30A;2D;480;480;0.5",
"Dolomite, Crushed;90C36;30B;2D;1280;1600;2",
"Dolomite, Lumpy;95D36;30B;2D;1440;1600;2",
"Earth, Loam, Dry, Loose;76C36;30B;2D;1220;1220;1.2",
"Ebonite, Crushed;67C35;30A;1A,1B,1C;1010;1120;0.8",
"Egg Powder;16A35MPY;30A;1B;260;260;1",
"Epsom Salts (Magnesium Sulfate);45A35U;30A;1A,1B,1C;640;800;0.8",
"Feldspar, Ground;73A37;15;2D;1040;1280;2",
"Feldspar, Lumps;95D37;15;2D;1440;1600;2",
"Feldspar, Powder;100A36;30B;2D;1600;1600;2",
"Felspar, Screenings;78C37;15;2D;1200;1280;2",
"Ferrous Sulfide, 1/2 inch (Iron Sulfide, Pyrites);128C26;30B;1A,1B,1C;1920;2160;2",
"Ferrous Sulfide, 100M (Iron Sulfide, Pyrites);113A36;30B;1A,1B,1C;1680;1920;2",
"Ferrous Sulphate (Iron Sulphate, Copperas);63C35U;30A;2D;800;1200;1",
"Filter-Aid (Diatomaceous Earth, Precoat);14A36Y;30B;3D;180;270;1.6",
"Fish Meal;38C45HP;30A;1A,1B,1C;560;640;1",
"Fish Scrap;45D45H;30A;2A,2B,2C;640;800;1.5",
"Flaxseed;44B35X;30A;1A,1B,1C;690;720;0.4",
"Flaxseed Cake (Linseed Cake);49D45W;30A;2A,2B;770;800;0.7",
"Flaxseed Meal (Linseed Meal);35B45W;30A;1A,1B;400;720;0.4",
"Flour Wheat;37A45LP;30A;1B;530;640;0.6",
"Flue Dust, Basic Oxygen Furnace;53A36LM;30B;3D;720;960;3.5",
"Flue Dust, Blast Furnace;118A36;30B;3D;1760;2000;3.5",
"Flue Dust, Boiler H. Dry;38A36LM;30B;3D;480;720;2",
"Fluorspar, Fine (Calcium Floride);90B36;30B;2D;1280;1600;2",
"Fluorspar, Lumps;100D36;30B;2D;1440;1760;2",
"Flyash;38A36M;30B;3D;480;720;2",
"Foundry Sand, Dry (Sand);95D37Z;15;3D;1440;1600;2.6",
"Fuller™s Earth, Calcined;40A25;45;3D;640;640;2",
"Fuller™s Earth, Dry, Raw (Bleach Clay);35A25;45;2D;480;640;2",
"Fuller™s Earth, Oily, Spent (Spent Bleach Clay);63C45OW;30A;3D;960;1040;2",
"Galena (Lead Sulfide);250A35R;30A;2D;3840;4160;5",
"Gelatine, Granulated;32B35PU;30A;1B;510;510;0.8",
"Gilsonite;37C35;30A;3D;590;590;1.5",
"Glass, Batch;90C37;15;3D;1280;1600;2.5", "Glue, Ground;40B45U;30A;2D;640;640;1.7",
"Glue, Pearl;40C35U;30A;1A,1B,1C;640;640;0.5",
"Glue, Veg. Powdered;40A45U;30A;1A,1B,1C;640;640;0.6",
"Gluten, Meal (Dry Corn Gluten);40B35P;30A;1B;640;640;0.6",
"Gluten, Meal (Dry Corn Gluten) Syral;40B35P;30A;1B;500;500;0.6",
"Gluten, Meal (Wet Corn Gluten);43B35OPH;30A;1B;690;690;2.2",
"Granite, Fine;85C27;15;3D;1280;1440;2.5",
"Grape, Pomace;18D45U;30A;2D;240;320;1.4",
"Graphite Flake (Plumago);40B25LP;45;1A,1B,1C;640;640;0.5",
"Graphite Flour;28A35LMP;30A;1A,1B,1C;450;450;0.5",
"Graphite Ore;70D35L;30A;2D;1040;1200;1",
"Guano Dry**;70C35;30A;3A,3B;1120;1120;2",
"Gypsum, Calcined (Plaster of Paris);58B35U;30A;2D;880;960;1.6",
"Gypsum, Calcined, Powdered (Plaster of Paris);70A35U;30A;2D;960;1280;2",
"Gypsum, Raw 1 inch(Calc.Sulfate, Plast.of Paris);75D25;30A;2D;1120;1280;2",
"Hay, Chopped **;10C35JY;30A;2A, 2B;130;190;1.6",
"Hexanedioic Acid (Adipic Acid);45A35;30A;2B;720;720;0.5",
"Hisarna Granulaat;--;--;--,--,--;--;2000;4",                       'Toegevoegd 25-6-2016
"Hominy, Dry;43C25D;30A;1A,1B,1C;560;800;0.4",
"Hops, Spend, Dry;35D35;30A;2A,2B,2C;560;560;1",
"Hops, Spent, Wet;53D45V;30A;2A,2B;800;880;1.5",
"Ice, Crushed;40D35O;30A;2A,2B;560;720;0.4",
"Ice, Cubes;34D35O;30A;1B;530;560;0.4",
"Ice, Flaked**;43C35O;30A;1B;640;720;0.6",
"Ice, Shell;34D45O;30A;1B;530;560;0.4",
"Ilmenite Ore (Titanium Dioxide);150D37;15;3D;2240;2560;2",
"Iron Ore Concentrate;150A37;15;3D;1920;2880;2.2",
"Iron Oxide Pigment;25A36LMP;30B;1A,1B,1C;400;400;1",
"Iron Oxide, Millscale;75C36;30B;2D;1200;1200;1.6",
"Iron Sulphate (Ferrous Sulfate);63C35U;30A;2D;800;1200;1",
"Iron Vitriol (Ferrous Sulfate);63C35U;30A;2D;800;1200;1",
"Kafir (Corn);43C25;45;3D;640;720;0.5",
"Kaolin Clay;63D25;30A;2D;1010;1010;2",
"Kaolin Clay (Tale);49A35LMP;30A;2D;670;900;2",
"Lactose;32A35PU;30A;1B;510;510;0.6",
"Lead Arsenate;72A35R;30A;1A,1B,1C;1150;1150;1.4",
"Lead Carbonate (Cerrusite);250A35R;30A;2D;3840;4160;1",
"Lead Ore, 1/2 inch;205C36;30B;3D;2880;3680;1.4",
"Lead Ore, 1/8 inch;235B35;30A;3D;3200;4320;1.4",
"Lead Oxide (Red Lead, Litharge) 100 Mesh;90A35P;30A;2D;480;2400;1.2",
"Lead Oxide (Red Lead, Litharge) 200 Mesh;105A35LP;30A;2D;480;2880;1.2",
"Lignite (Coal Lignite);41D35T;30A;2D;590;720;1",
"Limanite, Ore, Brown;120C47;15;3D;1920;1920;1.7",
"Lime Hydrated (Calcium Hydrate, Hydroxide);40B35LM;30A;2D;640;640;0.8",
"Lime Pebble;55C25HU;45;2A,2B;850;900;2",
"Lime, Ground, Unslaked (Quicklime);63B35U;30A;1A,1B,1C;960;1040;0.6",
"Lime, Hydrated, Pulverized;36A35LM;30A;1A,1B;510;640;0.6",
"Limestone, Agricultural (Calcium Carbonate);68B35;30A;2D;1090;1090;2",
"Limestone, Crushed (Calcium Carbonate); 88D36;30B;2D;1360;1440;2",
"Limestone, Dust (Calcium Carbonate);75A46MY;30B;2D;880;1520;1.8",
"Lindane (Benzene Hexachloride);56A45R;30A;1A,1B,1C;900;900;0.6",
"Linseed (Flaxseed);44B35X;30A;1A,1B,1C;690;720;0.4",
"Lithopone;48A35MR;30A;1A,1B;720;800;1",
"Magnesium Chloride (Magnesite);33C45;30A;1A,1B,1C;530;530;1",
"Maize (Milo);43B15N;45;1A,1B,1C;640;720;0.4",
"Malt, Dry Whole;25C35N;30A;1A,1B,1C;320;480;0.5",
"Malt, Dry, Ground;25C35N;30A;1A,1B,1C;320;480;0.5",
"Malt, Meal;38B25P;30A;1A,1B,1C;580;640;0.4",
"Malt, Sprouts;14C35P;30A;1A,1B,1C;210;240;0.4",
"Manganese Dioxide**;78A35NRT;30A;2A,2B;1120;1360;1.5",
"Manganese Ore;133D37;15;3D;2000;2240;2",
"Manganese Oxide;120A36;30B;2D;1920;1920;2",
"Manganese Sulfate;70C37;15;3D;1120;1120;2.4",
"Marble, Crushed;88B37;15;3D;1280;1520;2",
"Marl (Clay);80D36;30B;2D;1280;1280;1.6",
"Meat, Ground;53E45HQTX;30A;2A;800;880;1.5",
"Meat, Scrap (W/bone);40E46H;30B;2B;640;640;1.5",
"Mica, Flakes;20B16MY;30B;2D;270;350;1",
"Mica, Ground;14B36;30B;2D;210;240;0.9",
"Mica, Pulverized;14A36M;30B;2D;210;240;1",
"Milk, Dried, Flake;6B35PUY;30A;1B;80;100;0.4",
"Milk, Malted;29A45PX;30A;1B;430;480;0.9",
"Milk, Powdered;33B25PM;45;1B;320;720;0.5",
"Milk, Sugar;32A35PX;30A;1B;510;510;0.6",
"Milk, Whole, Powdered;28B35PUX;30A;1B;320;580;0.5",
"Mill Scale (Steel);123E46T;30B;3D;1920;2000;3",
"Milo Maize (Kafir);43B15N;45;1A,1B,1C;640;720;0.4",
"Milo, Ground (Sorghum Seed, Kafir);34B25;45;1A,1B,1C;510;580;0.5",
"Molybdenite Powder;107B26;30B;2D;1710;1710;1.5",
"Motar, Wet**;150E46T;30B;3D;2400;2400;3",
"Mustard Seed;45B15N;45;1A,1B,1C;720;720;0.4",
"Naphthalene Flakes;45B35;30A;1A,1B,1C;720;720;0.7",
"Niacin (Nicotinic Acid);35A35P;30A;2D;560;560;0.8",
"Oat Hulls;10B35NY;30A;1A,1B,1C;130;190;0.5",
"Oats;26C25MN;45;1A,1B,1C;420;420;0.4",
"Oats, Crimped;23C35;30A;1A,1B,1C;300;420;0.5",
"Oats, Crushed;22B45NY;30A;1A,1B,1C;350;350;0.6",
"Oats, Flour;35A35;30A;1A,1B,1C;560;560;0.5",
"Oats, Rolled;22C35NY;30A;1A,1B,1C;300;380;0.6",
"Oleo (Margarine);59E45HKPWX;30A;2A,2B;950;950;0.4",
"Orange Peel, Dry;1.5E+46;30A;2A,2B;240;240;1.5",
"Oyster Shells, Ground;55C36T;30B;3D;800;960;1.8",
"Oyster Shells, Whole;80D36TV;30B;3D;1280;1280;2.3",
"Paper Pulp (4% Or less);6.2E+46;30A;2A,2B;990;990;1.5",
"Paper Pulp (6% to 15%);6.2E+46;30A;2A,2B;960;990;1.5",
"Paraffin Cake, 1/2 inch;45C45K;30A;1A,1B;720;720;0.6",
"Peanut Meal;30B35P;30A;1B;480;480;0.6",
"Peanuts, Clean, in shell;18D35Q;30A;2A,2B;240;320;0.6",
"Peanuts, Raw (Uncleaned, Unshelled);18D36Q;30B;3D;240;320;0.7",
"Peanuts, Shelled;40C35Q;30A;1B;560;720;0.4",
"Peas, Dried;48C15NQ;45;1A,1B,1C;720;800;0.5",
"Perlite, Expanded;10C36;30B;2D;130;190;0.6",
"Phosphate Acid Fertilizer;60B25T;45;2A,2B;960;960;1.4",
"Phosphate Disodium (Sodium Phosphate);55A35;30A;1A,1B;800;960;0.9",
"Phosphate Rock, Broken;80D36;30B;2D;1200;1360;2.1",
"Phosphate Rock, Pulverized;60B36;30B;2D;960;960;1.7",
"Phosphate Sand;95B37;15;3D;1440;1600;2",
"Polyethylene, Resin Pellets;33C45Q;30A;1A,1B;480;560;0.4",
"Polystyrene Beads;40B35PQ;30A;1B;640;640;0.4",
"Polyvinyl Chloride Powder (PVC);25A45KT;30A;2B;320;480;1",
"Polyvinyl, Chloride Pellets;25E45KPQT;30A;1B;320;480;0.6",
"Potash (Muriate) Dry;70B37;15;3D;1120;1120;2",
"Potash (Muriate) Mine Run;75D37;15;3D;1200;1200;2.2",
"Potassium Carbonate;51B36;30B;2D;820;820;1",
"Potassium Nitrate, 1/2 inch (Saltpeter);76C16NT;30B;3D;1220;1220;1.2",
"Potassium Nitrate, 1/8 inch (Saltpeter);80B26NT;30B;3D;1280;1280;1.2",
"Potassium Sulfate;45B46X;30B;2D;670;770;1",
"Potassium-Chloride Pellets;125C25TU;45;3D;1920;2080;1.6",
"Potato Flour;48A35MNP;30A;1A,1B;770;770;0.5",
"PTA Crystal Slurry;VTK;--;--,--;--;1100;2.0",             'Toegevoegd 18-02-2016
"Pumice, 1/8 inch;45B46;30B;3D;670;770;1.6",
"Pyrite, Pellets;125C26;30B;3D;1920;2080;2",
"Quartz, 1/2 inch (Silicon Dioxide);85C27;15;3D;1280;1440;2",
"Quartz,100 Mesh (Silicon Dioxide);75A27;15;3D;1120;1280;1.7",
"Rape Seed Meal (Canola);38;?;?;540;660;0.8",
"Rice, Bran;20B35NY;30A;1A,1B,1C;320;320;0.4",
"Rice, Grits;44B35P;30A;1A,1B,1C;670;720;0.4",
"Rice, Hulled;47C25P;45;1A,1B,1C;720;780;0.4",
"Rice, Hulls;21B35NY;30A;1A,1B,1C;320;340;0.4",
"Rice, Polished;30C15P;45;1A,1B,1C;480;480;0.4",
"Rice, Rough;34C35N;30A;1A,1B,1C;510;580;0.6",
"Rosin, 1/2 inch;67C45Q;30A;1A,1B,1C;1040;1090;1.5",
"Rubber, Pelleted;53D45;30A;2A,2B,2C;800;880;1.5",
"Rubber, Reclaimed Ground;37C45;30A;1A,1B,1C;370;800;0.8",
"Rye;45B15N;45;1A,1B,1C;670;770;0.4",
"Rye Bran;18B35Y;30A;1A,1B,1C;240;320;0.4",
"Rye Feed;33B35N;30A;1A,1B,1C;530;530;0.5",
"Rye Meal;38B35;30A;1A,1B,1C;560;640;0.5",
"Rye Middlings;42B35;30A;1A,1B;670;670;0.5",
"Rye, Shorts;33C35;30A;2A,2B;510;530;0.5",
"Safflower Seed (Saffron);45B15N;45;1A,1B,1C;720;720;0.4",
"Safflower, Cake (Saffron);50D26;30B;2D;800;800;0.6",
"Safflower, Meal (Saffron);50B35;30A;1A,1B,1C;800;800;0.6",
"Sal Ammoniac (Ammonium Chloride);49A45FRS;30A;1A,1B,1C;720;830;0.7",
"Salicylic Acid;29B37U;15;3D;460;460;0.6",
"Salt Cake, Dry Coarse (Sodium Sulfate);85B36TU;30B;3D;1360;1360;2.1",
"Salt Cake, Dry Pulverized (Sodium Sulfate);75B36TU;30B;3D;1040;1360;1.7",
"Salt, Dry Coarse (Sodium Chloride);53C36TU;30B;3D;720;960;1",
"Salt, Dry Fine (Sodium Chloride);75B36TU;30B;3D;1120;1280;1.7",
"Sand (Resin Coated) Silica;104B27;15;3D;1670;1670;2",
"Sand (Resin Coated) Zircon;115A27;15;3D;1840;1840;2.3",
"Sand Dry Bank (Damp);120B47;15;3D;1760;2080;2.8",
"Sand Dry Bank (Dry);100B37;15;3D;1440;1760;1.7",
"Sand Dry Silica;95B27;15;3D;1440;1600;2",
"Sand Foundry (Shake Out);95D37Z;15;3D;1440;1600;2.6",
"Sawdust, Dry;12B45UX;30A;1A,1B,1C;160;210;0.7",
"Sea-Coal;65B36;30B;2D;1040;1040;1",
"Sesame Seed;34B26;30B;2D;430;660;0.6",
"Shale, Crushed;88C36;30B;2D;1360;1440;2",
"Shellac, Powdered or Granulated;31B35P;30A;1B;500;500;0.6",
"Silica Gel, 1/2 to 3 inch;45D37HKQU;15;3D;720;720;2",
"Silica, Flour;80A46;30B;2D;1280;1280;1.5",
"Slag, Blast Furnace Crushed;155D37Y;15;3D;2080;2880;2.4",
"Slag, Furnace Granular, Dry;63C37;15;3D;960;1040;2.2",
"Slate, Crushed, 1/2 inch;85C36;30B;2D;1280;1440;2",
"Slate, Ground, 1/8 inch;84B36;30B;2D;1310;1360;1.6",
"Sludge, Sewage, Dried;45E47TW;15;3D;640;800;0.8",
"Sludge, Sewage, Dry Ground;50B46S;30B;2D;720;880;0.8",
"Soap Detergent;33B35FQ;30A;1A,1B,1C;240;800;0.8",
"Soap, Beads or Granules;25B35Q;30A;1A,1B,1C;240;560;0.6",
"Soap, Chips;20C35Q;30A;1A,1B,1C;240;400;0.6",
"Soap, Flakes;10B35QXY;30A;1A,1B,1C;80;240;0.6",
"Soap, Powder;23B25X;45;1A,1B,1C;320;400;0.9",
"Soapstone, Talc, Fine;45A45XY;30A;1A,1B,1C;640;800;2",
"Soda Ash, Heavy (Sodium Carbonate);60B36;30B;2D;880;1040;1",
"Soda Ash, Light (Sodium Carbonate);28A36Y;30B;2D;320;560;0.8",
"Sodium Aluminate, Ground;72B36;30B;2D;1150;1150;1",
"Sodium Aluminum Sulphate**;75A36;30B;2D;1200;1200;1",
"Sodium Bicarbonate (Baking Soda);48A25;45;1B;640;880;0.6",
"Sodium Nitrate;75D25NS;30A;2A,2B;1120;1280;1.2",
"Sodium Phosphate;55A35;30A;1A,1B;800;960;0.9",
"Sodium Sulfite;96B46X;30B;2D;1540;1540;1.5",
"Soybean Meal Hot;40B35T;30A;2A,2B;640;640;0.5",
"Soybean Meal, Cold;40B35;30A;1A,1B,1C;640;640;0.5",
"Soybean, Cake;42D35W;30A;2A,1B,1C;640;690;1",
"Soybean, Cracked;35C36NW;30B;2D;480;640;0.5",
"Soybean, Flake, Extracted, Wet;34C35;30A;1A,1B,1C;540;540;0.8",
"Soybean, Flake, Raw;22C35Y;30A;1A,1B,1C;290;400;0.8",
"Soybean, Flour;29A35MN;30A;1A,1B,1C;430;480;0.8",
"Soybeans, Whole;48C26NW;30B;3D;720;800;1",
"Starch dry patato;38A15M;45;1A,1B,1C;400;500;1",
"Starch wet patato;38A15M;45;1A,1B,1C;400;750;1",
"Starch dry wheat;38A15M;45;1A,1B,1C;400;550;1",
"Steel Turnings, Crushed;125D46WV;30B;3D;1600;2400;3",
"Sugar Beet, Pulp, Dry;14C26;30B;2D;190;240;0.9",
"Suga Beet, Pulp, Wet;35C35X;30A;1A,1B,1C;400;720;1.2",
"Sugar, Powdered;55A35PX;30A;1B;800;960;0.8",
"Sugar, Raw;60B35PX;30A;1B;880;1040;1.5",
"Sugar, Refined, Granulated Dry;53B35PU;30A;1B;800;880;1.2",
"Sugar, Refined, Granulated Wet;60C35P;30A;1B;880;1040;2",
"Sulphur, Crushed, 1/2 inch;55C35N;30A;1A,1B;800;960;0.8",
"Sulphur, Lumpy, 3 inch;83D35N;30A;2A,2B;1280;1360;0.8",
"Sulphur, Powdered;55A35MN;30A;1A,1B;800;960;0.6",
"Sunflower Seed;29C15;45;1A,1B,1C;300;610;0.5",
"Sunflower Seed Flakes;28C35;30A;1A,1B,1C;430;450;0.8",
"Swee Bran Feed;29B45P;30A;1A,1B,1C;340;590;0.6",
"Talcum Powder;55A36M;30B;2D;800;960;0.8",
"Talcum, 1/2 ich;85C36;30B;2D;1280;1440;0.9",
"Tanbark, Ground**;55B45;30A;1A,1B,1C;880;880;0.7",
"Timothy Seed;36B35NY;30A;1A,1B,1C;580;580;0.6",
"Titanium Dioxide based pigments (powder);42C36FLO;15;3D;540;800;2",
"Tobacco, Scraps;20D45Y;30A;2A,2B;240;400;0.8",
"Tobacco, Snuff;30B45MQ;30A;1A,1B,1C;480;480;0.9",
"Tricalcium Phosphate;45A45;30A;1A,1B;640;800;1.6",
"Triple Sugar Phosphate ;53B36RS;30B;3D;800;880;2",
"Trisodium Phosphate;60C36;30B;2D;960;960;1.7",
"Trisodium Phosphate Granular;60B36;30B;2D;960;960;1.7",
"Trisodium Phosphate, Pulverized;50A36;30B;2D;800;800;1.6",
"Tung Nut Meats, Crushed;28D25W;30A;2A,2B;450;450;0.8",
"Tung Nuts ;28D15;30A;2A,2B;400;480;0.7",
"Urea Prills, Coated;45B25;45;1A,1B,1C;690;740;1.2",
"Vermiculite, Expanded;16C35Y;30A;1A,1B;260;260;0.5",
"Vermiculite, Ore;80D36;30B;2D;1280;1280;1",
"Vetch;48B16N;30B;1A,1B,1C;770;770;0.4",
"Walnut Shells, Crushed;40B36;30B;2D;560;720;1",
"Wheat; 47C25N;45;1A,1B,1C;720;770;0.4",
"Wheat Flour;37A45LP;45;1B;530;640;0.6",
"Wheat, Cracked;43B25N;45;1A,1B,1C;640;720;0.4",
"Wheat, Germ;23B25;45;1A,1B,1C;290;450;0.4",
"White Lead, Dry;88A36MR;30B;2D;1200;1600;1",
"Wood Chips, Screened;20D45VY;30A;2A,2B;160;480;0.6",
"Wood Flour;26B35N;30A;1A,1B;260;580;0.4",
"Wood Shavings;12E45VY;30A;2A,2B;130;260;1.5",
"Zinc Oxide, Heavy;33A45X;30A;1A,1B;480;560;1",
"Zinc Oxide, Light;13A45XY;30A;1A,1B;160;240;1",
"Zinc, Concentrate Residue;78B37;15;3D;1200;1280;1"}


    '--- "Oude benaming;Norm:;EN10027-1;Werkstof;[mm/m1/100°C];Poisson ;kg/m3;E [Gpa];Rm (20c);Rp0.2(0c);Rp0.2(20c);Rp(50c);Rp(100c);Rp(150c);Rp(200c);Rp(250c);Rp(300c);Rp(350c);Rp(400c);Equiv-ASTM;Opmerking",
    Public Shared steel() As String =
     {"Oude benaming;Norm:;EN10027-1;Werkstof;[mm/m1/100°C];Poisson ;kg/m3;E [Gpa];Rm (20c);0;20;50;100;150;Rp(200c);Rp(250c);Rp(300c);Rp(350c);Rp(400c);Equiv-ASTM;Opmerking",
    "Domex 690XPD(E);EN10149-2 UNS;S700MCD(E);1.8974;1.29;0.28;7850;192;810;675;740;765;690;675;660;640;620;580;540;--;--",
    "Duplex(Avesta-2205);EN 10088-1 UfllW;X2CrNiMoN22-5-3 saisna;1.4462;1.4;0.28;7800;200;640-950;335;460;385;360;335;315;300;0;0;0;A240-S31803;Max 300c",
    "Hastelloy-C22;DIN Nr: ASTM UNS;NiCr21Mo14W 2277 B575 N06022;2.4602;1.25;0.29;9000;205;786-800;310;370;354;338;310;283;260;248;244;241;--;--",
    "Inconel- 600;DIN Nicrofer7216 ASTM SO ;NiCr15Fe Alloy 600 B168 NiCr15Fe8 Npsepo;2.4816;1.44;0.29;8400;214;550;170;240;185;180;170;165;160;155;152;150;--;--",
    "P265GH;EN 10028-2 UNS;P265GH ;1.0425;1.29;0.28;7850;192;410-530;205;255;234;215;205;195;175;155;140;130;A516-Gr60;--",
    "S235JRG2;EN 10025 UNS;S235JRG2 ;1.0038;1.29;0.28;7850;192;340-470;180;195;200;190;180;170;150;130;120;110;A283-GrC;--",
    "S355J2G3;EN10025 UNS;S355J2G3;1.057;1.29;0.28;7850;192;490-630;284;315;340;304;284;255;226;206;0;0;A299;Max 300c	",
    "SS 304;EN10088-2;X5CrNI18-10 S30400;1.4301;1.76;0.28;7900;200;520-750;142;210;165;157;142;127;118;110;104;98;A240-304;--",
    "SS 304L;EN10088-2;X2CrNi19-11 S30403;1.4306;1.76;0.28;7900;200;520-670;132;200;155;147;132;118;108;100;94;89;A240-304L;--",
    "SS 316;EN10088-2;X5CrNiMo17-12-2 S31600;1.4401;1.76;0.28;8000;200;520-680;162;220;180;177;162;147;137;127;120;115;A240-316;--",
    "SS 316L;EN10088-2;X2CrNiMo17-12-2 S31603;1.4404;1.76;0.28;8000;200;520-680;152;220;170;166;152;137;127;118;113;108;A240-316L;--",
    "SS 316TI;EN10088-2;X6CrNiMoTi17-12-2 S31635;1.4571;1.76;0.28;8000;200;520-690;177;220;191;185;177;167;157;145;140;135;A240-316Ti;--",
    "SS 321;EN10088-2;X6CrNiTi18-10 S32100;1.4541;1.76;0.28;7900;200;500-720;167;200;184;176;167;157;147;136;130;125;A240-321;--",
    "SS 410 ;EN 10088-1 U1S;X12Cr13 (Gegloeid) 541000;1.4006;1.15;0.28;7700;216;450-650;230;250;240;235;230;225;225;220;210;195;A240-410;--",
    "SuperDuplex;--;X2CrNiMoN22-5-3 saisna;1.4501;1.4;0.28;7800;200;730-930;445;550;510;480;445;405;400;395;0;0;--;--"}

    'DN, inch,OD, wall1, wall2, wall3,...
    Public Shared pipe_ss() As String =
   {"DN100;4 inch; 114.3;  6.3;7.1;8;10;12.7;16.0",
    "DN125;5 inch; 139.7;  6.3;7.1;8;10;12.7;16.0",
    "DN150;6 inch; 168.3;  6.3;7.1;8;10;12.7;16.0",
    "DN200;8 inch; 219.1;  6.3;7.1;8;10;12.7;16.0",
    "DN250;10 inch; 273;   6.3;7.1;8;10;12.7;16.0",
    "DN300;12 inch; 323.9; 6.3;7.1;8;10;12.7;16.0",
    "DN350;14 inch; 355.6; 6.3;7.1;8;10;12.7;16.0",
    "DN400;16 inch; 406.4; 6.3;7.1;8;10;12.7;16.0",
    "DN500;20 inch; 508;   6.3;7.1;8;10;12.7;16.0"}

    Public Shared pipe_steel() As String =
   {"DN100;4 inch; 114.3;  6.02;  8.56; 0;   0",
    "DN125;5 inch; 141.3;  6.55;  9.53; 0;   0",
    "DN150;6 inch; 168.3;  7.11; 10.97; 0;   0",
    "DN200;8 inch; 219.1;  6.35;  8.18; 12.7;0",
    "Specl;.. inch; 229.0;  20.00;  20.00; 20.00;0",
    "DN250;10 inch; 273;   6.35;  9.27; 12.7;0",
    "DN300;12 inch; 323.9; 6.35;  9.27; 12.7;0",
    "DN350;14 inch; 355.6; 7.92;  9.53; 0;   0",
    "DN400;16 inch; 406.4; 7.92;  9.53; 0;   0"}


    Public Shared motorred() As String =
     {"Description; Speed; power;cost;shaftdia",
     "0.18 Kw,R27DR63M4;69.5;0.18;253.51;25",
     "3 Kw, Bauer BG60-11/DHE11XAC-TF;49.5;3;1132;50",
     "3 Kw, 20rpmR107;20;3;1908.74;70",
     "3 Kw, R77DRM100L4;29.12;3;896.25;40",
     "3 Kw, R137R77/II2GD EDRE100LC4;6.2;3;3851.01;90",
     "1.1 Kw, R77/II2GD EDRE90M4;27;1.1;814.50;40",
     "0.75 Kw, R87DRE90L6;940;0.75;1003.71;50",
     "2.2 Kw, R47DRE100M4;14.56;2.2;471.18;30",
      "1.1 Kw, R97DRN90S4;186;1.1;1340.06;60",
      "2.2 Kw 15.1023;55;2.2;0;0",
      "1.5 Kw, 1440 rpm;1440;1.5;0;0",
      "2.2 Kw, 1440 rpm;1440;2.2;0;0",
      "3 Kw, 1455 rpm;1455;3;0;0",
      "4 Kw, 1465 rpm;1465;4;0;0",
      "5.5 Kw, 1475 rpm;1475;5.5;0;0",
      "7.5 Kw, 1475 rpm;1475;7.5;0;0",
      "9.2 Kw, 1475 rpm;1475;9.2;0;0",
      "11 Kw, 1475 rpm;1475;11;0;0",
      "15 Kw, 1475 rpm;1475;15;0;0",
      "18.5 Kw, 1480 rpm;1480;18.5;0;0",
      "22 Kw, 1482 rpm;1482;22;0;0",
      "30 Kw, R137DRP225S4/TF/PT;53;30;5802.48;90",
      "45 Kw, R147DRP250M4/TF/NIB/PT;49;45;8588.74;110",
      "37 Kw, 1482 rpm;1482;37;0;0"}

    Public Shared coupl() As String =
     {"Diameter;cost,percentage na korting",
      "58 mm, n-eupexB;102.7;0.55",
      "68 mm, n-eupexB;111.24;0.55",
      "80 mm, n-eupexB;127.5;0.55",
      "95 mm, n-eupexB;159.58;0.55",
      "160 mm, n-eupexB;294.50;0.45",
      "flender, n-eupexB180;264.94;1",
      "flender, n-eupexB225;313.77;1",
      "N-eupexA250;509.57;1",
      "flender, n-eupexB250;387.75;1",
      "flender, n-eupexA280;763.26;1",
      "flender, n-eupexA315;852.50;1",
      "flender, n-eupexA400;1545.50;1",
      "200 mm, n-eupexB;480;0.45",
      "110 mm, n-eupexA;291.6;0.55",
      "125 mm, n-eupexA;407.35;0.55",
      "140 mm, n-eupexA;534.4;0.55",
      "160 mm, n-eupexA;742.05;0.55",
      "180 mm, n-eupexA;904.4;0.55",
      "200 mm, n-eupexA;1173.5;0.55",
      "225 mm, n-eupexA;1580;0.55",
      "250 mm, n-eupexA;1907.1;0.55",
      "280 mm, n-eupexA;2339;0.55",
      "315 mm, n-eupexA;3190.7;0.55",
      "350 mm, n-eupexA;4433;0.55",
      "400 mm, n-eupexA;5778;0.55",
      "440 mm, n-eupexA;7167.5;0.55",
      "480 mm, n-eupexA;8983.5;0.55"}


    Public Shared ppaint() As String =
     {"Description;cost",
      "Pickling + passivating; 0.50",
      "10-20m2 75um zink compound;13.25",
      "20-100m2 75um zink compound;12.50",
      "10-20m2 150um primer en epoxy (binnen);17.0",
      "20-100m2 150um primer en epoxy (binnen);16.40",
      "10-20m2 150um primer en polyurethaan (buiten);17.90",
      "20-100m2 150um primer en polyurethaan (buiten);17.25",
      "10-20m2 225um primer, midcoat, polyurethaan (buiten);18.60",
      "20-100m2 225um primer, midcoat, polyurethaan (buiten);18.15",
      "10-20m2 330um primer, midcoat, polyurethaan (buiten);20.75",
      "20-100m2 330um primer, midcoat, polyurethaan (buiten);20.0",
      "10-20m2 75um primer, zincsilicaat -90C/+400C;13.25",
      "20-100m2 75um primer,  zincsilicaat -90C tot +400C;12.50",
      "10-20m2 120um primer,  zincsilicaat -90C tot +400C;19.50",
      "20-100m2 120um primer,  zincsilicaat -90C tot +400C;18.0",
      "10-20m2 250um primer, midcoat, polyurethaan;23.0",
      "20-100m2 250um primer, midcoat, polyurethaan;22.0"}


    Public Shared lager() As String = 'T=trekbus, C=cylindrisch, zie SKFboekje
     {"diameter;prijs",
     "40 mm Trekbus;120.45",
     "50 mm Trekbus;152.57",
      "55 mm Trekbus;170",
      "60 mm Trekbus;196.23",
      "65 mm Trekbus;250",
      "70 mm Trekbus;311.98",
      "75 mm Trekbus;315",
      "80 mm Trekbus;318.59",
      "85 mm Trekbus;370",
      "90 mm Trekbus;420.29",
      "100 mm Trekbus;553.54",
      "110 mm Trekbus;590.63",
      "115 mm Trekbus;729.07",
      "120 mm Trekbus;800",
      "125 mm Trekbus;901.85",
      "130 mm Trekbus;1000",
      "135 mm Trekbus;1111.90",
      "140 mm Trekbus;1385.19",
      "180 mm Trekbus;2505.50",
      "40 mm Cylindrisch;94.39",
      "50 mm Cylindrisch;106.76",
      "55 mm Cylindrisch;120",
      "60 mm Cylindrisch;142.49",
      "65 mm Cylindrisch;170",
      "70 mm Cylindrisch;260.12",
      "75 mm Cylindrisch;270",
      "80 mm Cylindrisch;280.69",
      "85 mm Cylindrisch;290",
      "90 mm Cylindrisch;298.29",
      "95 mm Cylindrisch;340",
      "100 mm Cylindrisch;395.47",
      "110 mm Cylindrisch;526.52",
      "120 mm Cylindrisch;629.66",
      "130 mm Cylindrisch;777.68",
      "140 mm Cylindrisch;962.66",
      "150 mm Cylindrisch;1187.03",
      "160 mm Cylindrisch;1474.77",
      "210 mm Cylindrisch;1700",
      "360 mm Cylindrisch;3500"}

    Public Shared astap_dia() As String =   'tbv as diameter selectie
      {"Dia;empty",
      "100 ;0",
      "110 ;0",
      "120 ;0",
      "140 ;0",
      "160 ;0",
      "200 ;0",
      "210 ;0",
      "260 ;0",
      "310 ;0",
      "350 ;0",
      "400 ;0",
      "500 ;0"}

    Public Shared Flight_dia() As String =   'tbv screw diameter selectie
      {"Flight_dia;empty",
      "280 ;0",
      "330 ;0",
      "400 ;0",
      "500 ;0",
      "630 ;0",
      "800 ;0",
      "1000 ;0",
      "1200 ;0",
      "1400 ;0"}

    Public Shared pakking() As String =
     {"Merk;maat;prijs",
      "Flowtite wit, 3*1.5  ;53",
      "Flowtite wit, 5*2    ;73",
      "Flowtite wit, 7*2.5  ;83",
      "Flowtite wit, 10*3   ;103",
      "230x133x6 Silicone wit;8.15"}


    Public Shared emotor() As String = {"3.0; 1500", "4.0; 1500", "5.5; 1500", "7.5; 1500", "11;  1500", "15; 1500", "22; 1500",
                                       "30  ; 1500", "37;  1500", "45;  1500", "55;  1500", "75; 1500", "90; 1500",
                                       "110 ; 1500", "132; 1500", "160; 1500", "200; 1500"}


    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")


        For hh = 0 To (UBound(_inputs) - 1)              'Fill combobox1
            words = _inputs(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 225                   'Grafite ore

        '-------Fill combobox2, Steel selection------------------
        For hh = 0 To (UBound(steel) - 1)               'Fill combobox 2 with steel data
            words = steel(hh).Split(CType(";", Char()))
            ComboBox2.Items.Add(words(0))
        Next hh
        ComboBox2.SelectedIndex = 7                     'S355

        '-------Fill combobox5, emotor selection------------------
        For hh = 0 To (UBound(emotor) - 1)               'Fill combobox 5 emotor data
            words = emotor(hh).Split(CType(";", Char()))
            ComboBox5.Items.Add(words(0))
        Next hh
        ComboBox5.SelectedIndex = 0


        TextBox133.Text = "Plaat zwart" & vbTab & "1.30 €/kg" & vbCrLf
        TextBox133.Text &= "Plaat 304 " & vbTab & vbTab & "0.00 €/kg " & vbCrLf
        TextBox133.Text &= "Plaat 316 " & vbTab & vbTab & "0.00 €/kg " & vbCrLf
        TextBox133.Text &= vbCrLf
        TextBox133.Text &= "Pijp gelast" & vbCrLf
        TextBox133.Text &= "Pijp zwart  " & vbTab & vbTab & "5.30 €/kg " & vbCrLf
        TextBox133.Text &= "Pijp 304    " & vbTab & vbTab & "0.00 €/kg " & vbCrLf
        TextBox133.Text &= "Pijp 316    " & vbTab & vbTab & "0.00 €/kg " & vbCrLf
        TextBox133.Text &= vbCrLf
        TextBox133.Text &= "Bron Staalprijzen.nl" & vbCrLf

        Screw_combo_init()
        Pipe_dia_combo_init()
        Motorreductor()
        Coupling_combo()
        Lager_combo()
        Astap_combo()
        Paint_combo()
        Pakking_combo()
    End Sub
    Private Sub Calc_sequence()
        Calculate()
        Calulate_stress_1()
        Costing_material()
        Screen_contrast()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, Button1.Click, TabPage1.Enter, ComboBox11.SelectedValueChanged, NumericUpDown8.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Calculate()
        Dim cap_hr_100 As Double        '100% Capacity conveyor [m3/hr]
        Dim cap_hr_100_in As Double     '100% Capacity conveyor inclined [m3/hr]
        Dim actual_cap_m3 As Double     'actual Çapacity conveyor [m3/hr]
        Dim iso_forward As Double       'Power for forward motion
        Dim iso_incline As Double       'Power inclination
        Dim iso_no_product As Double    'Power for seals + bearings
        Dim iso_power As Double         'Total Power 
        Dim height As Double            'Height difference due to inclination 
        Dim mekog_pow As Double         'Mekog installed power
        Dim mekog_torque As Double      'Mekog installed torque
        Dim NON_torque As Double        'NON gearbox torque
        Dim NON_pow As Double           'NON gearbox power
        Dim flight_speed As Double      'Flight speed
        Dim r_time As Double
        Dim cap_under_angle As Double
        Dim filling_perc As Double      'Conveyor horizontal
        Dim filling_perc_incl As Double 'Conveyor horizontal

        '-------------- get data----------
        Double.TryParse(CType(ComboBox3.SelectedItem, String), _pipe_OD)
        _pipe_OD /= 1000                                    '[m]
        Label177.Text = _pipe_OD.ToString
        If ComboBox11.SelectedIndex > 0 Then
            Double.TryParse(CType(ComboBox11.SelectedItem, String), _diam_flight)
            _diam_flight /= 1000                            '[mm] -> [m]
            pitch = _diam_flight * NumericUpDown2.Value     '[m]
            _angle = NumericUpDown4.Value                   '[degree]
            speed = NumericUpDown7.Value                    '[rpm]
            progress_resistance = NumericUpDown9.Value      '[-]
            density = NumericUpDown6.Value                  '[kg/m3] Density
            _λ6 = NumericUpDown3.Value                      '[m] lengte van de trog/schroef 

            '------- Flight speed (ATEX < 1 [m/s])-----------
            flight_speed = speed / 60 * PI * _diam_flight   '[m/s]

            '------- Required Volumetric capacity --------
            _regu_flow_kg_hr = NumericUpDown5.Value * 1000  '[kg/hr] required flow
            actual_cap_m3 = _regu_flow_kg_hr / density      '[m3/hr] required flow

            '-------- Volumetric Capacity [m3/hr] ---------------
            '-------- Of the selected diameter ------------------
            cap_hr_100 = PI / 4 * (_diam_flight ^ 2 - _pipe_OD ^ 2) * pitch * speed * 60    ' [m3/hr]

            '-------- Inclination factor ------------------------
            cap_under_angle = -0.0213 * _angle + 1.0        'Inclination capacity factor
            cap_hr_100_in = cap_hr_100 * cap_under_angle    'capacity loss due to inclination 

            '--------------- now calc in [kg/hr] ---------------
            filling_perc = actual_cap_m3 / cap_hr_100 * 100             'Horizontal
            filling_perc_incl = actual_cap_m3 / cap_hr_100_in * 100     'Inclined

            Select Case RadioButton9.Checked
                Case True   'Transport screw
                    TextBox01.BackColor = CType(IIf(filling_perc > 45, Color.Red, Color.LightGreen), Color)
                Case False  'Metering screw
                    TextBox01.BackColor = CType(IIf(filling_perc > 75, Color.Red, Color.LightGreen), Color)
            End Select

            '--------------- ISO 7119 power calc -----------------
            height = _λ6 * Sin(_angle / 360 * 2 * PI)

            iso_forward = _regu_flow_kg_hr * _λ6 * 9.91 * progress_resistance / (3600 * 1000)    'Forwards [kW]
            iso_incline = _regu_flow_kg_hr * height * 9.81 / (3600 * 1000)                       'Uphill [kW]
            iso_no_product = _diam_flight * _λ6 / 20                                     'Power for seals 0. + bearings [kW]
            iso_power = iso_forward + iso_incline + iso_no_product

            '--------------- MEKOG power calc -----------------
            mekog_pow = Round(_regu_flow_kg_hr * _λ6 / (40 * 1.36 * 1000), 1)    '[kW]
            mekog_pow *= 1.6 'Based on current measurement Q19.1165 (Borouge 4) dd 12/09/2019
            mekog_torque = mekog_pow * 1000 / (speed * 2 * PI / 60)

            ' Debug.WriteLine("_regu_flow_kg_hr= " & _regu_flow_kg_hr.ToString)
            'Debug.WriteLine(" _λ6= " & _λ6.ToString)

            '--------------- NON asperen chart ----------------
            NON_torque = Calc_NON_Torque((_diam_flight * 1000), _λ6)    '[Nm]
            NON_pow = NON_torque * (speed * 2 * PI / 60) / 1000         '[kW]
            NON_pow /= 0.8                                              '[kW] (efficiency gearbox)

            '-------------- Retention time --------------------
            r_time = _λ6 / (speed / 60 * pitch)                         '[sec]

            '--------------- present results------------
            TextBox19.Text = _λ6.ToString
            TextBox11.Text = flight_speed.ToString("F2") 'Flight speed [m/s]
            TextBox18.Text = CType(_diam_flight * CDbl(1000.ToString), String)
            TextBox16.Text = CType(_pipe_OD * CDbl(1000.ToString), String)  'Pipe diameter [m]
            TextBox01.Text = filling_perc.ToString("F1")
            TextBox127.Text = filling_perc_incl.ToString("F1")  'Inclination factor
            TextBox03.Text = iso_power.ToString("F1")           '[kW]
            TextBox04.Text = mekog_pow.ToString("F1")           '[kW]
            TextBox137.Text = mekog_torque.ToString("F0")       '[Nm] gearbox
            TextBox139.Text = NON_pow.ToString("F1")            '[Nm] power
            TextBox138.Text = NON_torque.ToString("F0")         '[Nm] gearbox

            TextBox110.Text = r_time.ToString("F0")
            TextBox123.Text = cap_hr_100.ToString("F0")        '[m3/hr] @ 100% filling horizontal
            TextBox126.Text = cap_under_angle.ToString("F2")  'Inclination factor
            TextBox124.Text = actual_cap_m3.ToString("F1") '[m3/hr] 

            '--------------- checks ---------------------
            NumericUpDown7.BackColor = CType(IIf(speed > 45, Color.Red, Color.Yellow), Color)
            Label135.Visible = CBool(IIf(flight_speed > 1.0, True, False))
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabControl1.Enter, RadioButton8.CheckedChanged, RadioButton7.CheckedChanged, RadioButton6.CheckedChanged, RadioButton4.CheckedChanged, NumericUpDown35.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown25.ValueChanged, ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox13.SelectedIndexChanged, ComboBox12.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, CheckBox8.CheckedChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, CheckBox4.CheckedChanged, CheckBox7.CheckedChanged, CheckBox6.CheckedChanged, TabPage4.Enter, CheckBox5.CheckedChanged

        Calc_sequence()
    End Sub

    'Materiaal in de conveyor
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If NumericUpDown6.Value > -1 And NumericUpDown9.Value > -1 Then
            Dim words() As String = _inputs(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            NumericUpDown6.Value = CDec(words(5)) 'Density max
            NumericUpDown9.Value = CDec(words(6)) 'Material factor
            Label37.Text = "CEMA material code " & words(1)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage5.Enter, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown48.ValueChanged, TextBox99.VisibleChanged, RadioButton11.CheckedChanged, NumericUpDown50.ValueChanged, NumericUpDown49.ValueChanged
        Calc_sequence()
    End Sub
    'Please note complete calculation in [m] not [mm]
    Private Sub Calulate_stress_1()
        Dim qq As Double
        Dim q_load_1, q_load_2, q_load_comb, q_load_3, q As Double
        Dim force_1 As Double
        Dim force_2 As Double
        Dim force_3 As Double
        Dim Q_max_bend As Double
        Dim F_tangent, Radius_transport As Double
        Dim pipe_weight_m As Double
        Dim pipe_OR, pipe_IR As Double
        Dim sigma_eg As Double                      'Sigma eigen gewicht
        Dim flight_hoogte, flight_gewicht, flight_lengte_buiten, flight_lengte_binnen, flight_lengte_gem, fligh_dik As Double
        Dim P_torque, Tou_torque As Double           'Torque @ aandrijving
        Dim P_torque_M, Tou_torque_M As Double       'Torque @ max bend
        Dim words() As String
        Dim Ra, Rb, R_total As Double
        Dim Uniform_mat_load As Double
        Dim combined_stress As Double
        Dim max_sag As Double                       'maximale doorzakking pijp
        Dim xnul As Double                          'nul positie in dwarskrachtenlijm
        Dim Column_h(4) As Double                   'Material column height
        Dim Column_a(4) As Double                   'Inlet chute pipe area

        NumericUpDown13.Value = NumericUpDown7.Value

        If (ComboBox5.SelectedIndex > -1) Then      'Prevent exceptions
            words = emotor(ComboBox5.SelectedIndex).Split(CType(";", Char()))
            Double.TryParse(words(0), installed_power)
            Start_factor = NumericUpDown18.Value
            actual_power = installed_power * Start_factor
        End If

        If (ComboBox2.SelectedIndex > -1) Then      'Prevent exceptions
            words = steel(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            TextBox06.Text = words(6)     'Density steel

            '--------------- select the strength @ temperature
            qq = NumericUpDown11.Value
            Select Case True
                Case (qq > -10 AndAlso qq <= 0)
                    Double.TryParse(words(9), sigma02)     'Sigma 0.2 [N/mm]
                Case (qq > 0 AndAlso qq <= 20)
                    Double.TryParse(words(10), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 20 AndAlso qq <= 50)
                    Double.TryParse(words(11), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 50 AndAlso qq <= 100)
                    Double.TryParse(words(12), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 100 AndAlso qq <= 150)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 150 AndAlso qq <= 200)
                    Double.TryParse(words(13), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 200 AndAlso qq <= 250)
                    Double.TryParse(words(14), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 250 AndAlso qq <= 300)
                    Double.TryParse(words(15), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 300 AndAlso qq <= 350)
                    Double.TryParse(words(16), sigma02)    'Sigma 0.2 [N/mm]
                Case (qq > 350 AndAlso qq <= 400)
                    Double.TryParse(words(17), sigma02)    'Sigma 0.2 [N/mm]
            End Select
            TextBox07.Text = CType(sigma02, String)
            sigma_fatique = sigma02 * 0.3                   'Fatique stress uitgelegd op oneindige levensduur
            TextBox08.Text = sigma_fatique.ToString("F0")
        End If

        If ComboBox3.SelectedIndex > -1 Then
            words = pipe_steel(ComboBox3.SelectedIndex).Split(CType(";", Char()))

            Double.TryParse(words(2), _pipe_OD)
            _pipe_OD /= 1000                            'Outside Diameter [m]
            pipe_OR = _pipe_OD / 2                      'Radius [mm]
            _pipe_wall = CDbl(ComboBox6.SelectedItem)   'Wall thickness [mm]
            _pipe_wall /= 1000
            _pipe_ID = (_pipe_OD - 2 * _pipe_wall)      'Inside diameter [mm]
            pipe_IR = _pipe_ID / 2                      'Inside radius [mm]

            pipe_weight_m = PI / 4 * (_pipe_OD ^ 2 - _pipe_ID ^ 2) * 7850   'Weight per meter [kg/m]

            TextBox13.Text = pipe_weight_m.ToString("F1")                   'gewicht per meter
            TextBox16.Text = (_pipe_OD * 1000).ToString("F1")               'Diameter [m]

            '---------------- Traagheids moment Ix= PI/64.(D^4-d^4)---------------------
            pipe_Ix = PI / 64 * (_pipe_OD ^ 4 - _pipe_ID ^ 4)                  '[m4]
            TextBox26.Text = (pipe_Ix * 1000 ^ 4).ToString("F0")

            '---------------- Weerstand moment Buiging  Wx= PI/32.(D^4-d^4)/D---------------------
            pipe_Wx = PI / 32 * (_pipe_OD ^ 4 - _pipe_ID ^ 4) / _pipe_OD        '[m3]
            TextBox14.Text = (pipe_Wx * 1000 ^ 3).ToString("F0")

            '---------------- Weerstand moment Torsie (polair)  Wp= PI/16.(D^4-d^4)/D --------------
            pipe_Wp = PI / 16 * (_pipe_OD ^ 4 - _pipe_ID ^ 4) / _pipe_OD       '[m3]
            TextBox15.Text = (pipe_Wp * 1000 ^ 3).ToString("F0")

            '----------dimensions-----------------------------------------------------------------
            _κ1 = NumericUpDown31.Value                '[m] exposed pipe length force 1
            _κ2 = NumericUpDown32.Value                '[m] exposed pipe length force 2
            _κ3 = NumericUpDown22.Value                '[m] exposed pipe length force 3

            '---------- materiaal gewicht inlaat kolom op pipe--------------------------
            product_density = NumericUpDown6.Value      'product density [kg/m3]
            Column_h(0) = NumericUpDown17.Value         'Material height on top of pipe
            Column_a(0) = _pipe_OD * 1.0                '[m2] pipe area

            '-------- Inlet opening #1------
            Column_h(1) = NumericUpDown19.Value         '[m] column height
            Column_a(1) = _pipe_OD * _κ1                '[m2] pipe area

            '-------- Inlet opening #2------
            Column_h(2) = NumericUpDown36.Value         '[m] column height
            Column_a(2) = _pipe_OD * _κ2                '[m2] pipe area

            '-------- Inlet opening #3------
            Column_h(3) = NumericUpDown37.Value         '[m] column height
            Column_a(3) = _pipe_OD * _κ3               '[m2] pipe area

            Uniform_mat_load = Column_a(0) * Column_h(0) * product_density * 9.81         '[N/m]
            force_1 = Column_a(1) * Column_h(1) * product_density * 9.81    '[N]
            force_2 = Column_a(2) * Column_h(2) * product_density * 9.81    '[N]
            force_3 = Column_a(3) * Column_h(3) * product_density * 9.81    '[N]

            TextBox115.Text = force_1.ToString("0")             'Material inlet force
            TextBox116.Text = force_2.ToString("0")             'Material inlet force
            TextBox117.Text = force_3.ToString("0")             'Material inlet force
            TextBox118.Text = Uniform_mat_load.ToString("0")

            _λ1 = NumericUpDown39.Value            '[m] CL Bearing to drive plate
            _λ2 = NumericUpDown16.Value            '[m] CL inlet chute #1 to bearing
            _λ3 = NumericUpDown24.Value            '[m] CL inlet chute #2 to bearing
            _λ4 = NumericUpDown28.Value            '[m] CL inlet chute #3 to bearing
            _λ5 = NumericUpDown40.Value            '[m] CL Bearing to tail plate
            _λ6 = NumericUpDown3.Value             '[m] lengte van de trog/schroef 
            _λ7 = _λ1 + _λ6 + _λ5                  '[m] bearing-bearing

            '============= calc load ========================================
            '================================================================
            Young = NumericUpDown1.Value * 1000 '[N/mm2]
            '---------------- gewicht flight [mm] dik----------------------------------
            flight_hoogte = (_diam_flight - _pipe_OD / 1000) / 2                                '[m]
            flight_lengte_buiten = Sqrt((PI * _diam_flight) ^ 2 + (pitch) ^ 2)

            flight_lengte_binnen = Sqrt((PI * _pipe_OD / 1000) ^ 2 + (pitch) ^ 2)
            flight_lengte_gem = (flight_lengte_buiten + flight_lengte_binnen) / 2
            fligh_dik = NumericUpDown8.Value / 1000                                             '[m]
            flight_gewicht = (flight_hoogte * flight_lengte_gem * fligh_dik * 7850) / pitch     'Flight Gewicht per meter
            TextBox02.Text = flight_gewicht.ToString("F1")                                   'Flight Gewicht per meter
            TextBox05.Text = (fligh_dik * 1000).ToString("F1")                                 'Flight dikte [mm]

            '------------- aandrijving torsie @ drive ----------------------------
            P_torque = actual_power * 1000 / (2 * PI * NumericUpDown7.Value / 60)
            TextBox29.Text = P_torque.ToString("F0")                                        'Torque from drive [N.m]

            '----------- Weight (pipe+flight) + transport force combined ------
            '---- Worst case material assumed sitting lowest point of the trough---

            q_load_1 = (pipe_weight_m + flight_gewicht) * 9.81      '[N/m] Uniform distributed weight
            If CheckBox10.Checked Then q_load_1 = NumericUpDown29.Value 'Test
            TextBox22.Text = q_load_1.ToString("F0")                '[N/m] Uniform distributed weight

            '----------- Axial load caused by transport of product
            Radius_transport = (_diam_flight + _pipe_OD) / 4        'Acc Jos (D+d)/4
            F_tangent = P_torque / Radius_transport
            q_load_2 = F_tangent / _λ6                              'Transport kracht geeft doorbuiging pijp
            q_load_3 = Uniform_mat_load                             '[N/m] Uniform distributed material weight
            TextBox17.Text = q_load_3.ToString("F0")                '[N/m]

            '============= Traditionele VTK berekening ==========================
            '============= verwaarloosd Q_load2 =================================
            If CheckBox1.Checked Then
                q_load_2 = 0
            End If
            TextBox28.Text = q_load_2.ToString("F2")               '[N]
            q_load_comb = Sqrt((q_load_1 + q_load_3) ^ 2 + q_load_2 ^ 2)     '[N/m] Radiale en tangentiele kracht gecombineerd

            '============= Reactie krachten Bearings==============================
            '=====================================================================
            R_total = q_load_1 * _λ6        '[N] Steel weight (stub ends + pipe+flight)
            R_total += q_load_3 * _λ6       '[N] Material weight
            R_total += force_1              '[N] Material falling on the pipe chute #1
            R_total += force_2              '[N] Material falling on the pipe
            R_total += force_3              '[N] Material falling on the pipe

            'Momenten evenwicht om punt Ra, 
            'Moment= Kracht x arm
            'Gelijkmatigebelasting, Moment= Kracht * arm (arm= halve lengte) 
            Rb = (q_load_1 * _λ6) * (_λ1 + _λ6 * 0.5)  'Pipe weight
            Rb += (q_load_3 * _λ6) * (_λ1 + _λ6 * 0.5) 'Uniform load
            Rb += force_1 * _λ2             'Inlet force #1
            Rb += force_2 * _λ3             'Inlet force #2
            Rb += force_3 * _λ4             'Inlet force #3

            Rb /= _λ7
            Ra = R_total - Rb

            TextBox24.Text = Ra.ToString("F0")          'Reactie kracht Ra
            TextBox36.Text = Rb.ToString("F0")          'Reactie kracht Rb
            TextBox39.Text = R_total.ToString("F0") 'Reactie kracht Ra+Rb

            'Gebaseerd op http://beamguru.com/online/beam-calculator/
            '============ Krachten zijn
            TextBox86.Text = force_1.ToString("0")
            TextBox87.Text = force_2.ToString("0")
            TextBox89.Text = force_3.ToString("0")

            '=========== x posities gemeten vanaf de drive bearing=============
            TextBox7.Text = _λ2.ToString("0.0")   '[m] CL Inlaat #1-drive bearing
            TextBox8.Text = _λ3.ToString("0.0")   '[m] CL Inlaat #2-drive bearing
            TextBox9.Text = _λ4.ToString("0.0")   '[m] CL Inlaat #3-drive bearing

            'https://en.wikipedia.org/wiki/Direct_integration_of_a_beam#Sample_calculations
            '=========== Distance gemeten vanaf de inlaatschot=============
            Dim imax_count, i_chute_1, i_chute_2, i_chute_3 As Integer
            Dim ΔL As Double

            For i = 1 To _steps
                _d(i) = i / _steps * _λ7   'Chop conveyor in 100 pieces
            Next

            '=========== Shear Force vanaf de inlaatschot=============
            '=========== dwarskrachtenlijn (shear force) =============
            q = q_load_1 + q_load_3
            _s(0) = Ra
            ΔL = _λ7 / _steps

            For i = 1 To _steps
                If _d(i) > _λ1 And _d(i) < (_λ7 - _λ1) Then
                    _s(i) = _s(i - 1) - q * ΔL
                Else
                    _s(i) = _s(i - 1)
                End If
                If _d(i) > _λ2 - ΔL / 2 And _d(i) < _λ2 + ΔL / 2 Then
                    _s(i) -= force_1    '[N] shear force #1
                    i_chute_1 = i       'locatie
                End If
                If _d(i) > _λ3 - ΔL / 2 And _d(i) < _λ3 + ΔL / 2 Then
                    _s(i) -= force_2    '[N] shear force #2
                    i_chute_2 = i
                End If
                If _d(i) > _λ4 - ΔL / 2 And _d(i) < _λ4 + ΔL / 2 Then
                    _s(i) -= force_3    '[N] shear force #
                    i_chute_3 = i
                End If
            Next

            '=========== momentenlijn (bending moment )====================
            _m(0) = 0   'Simply supported
            For i = 1 To _steps
                _m(i) = _m(i - 1) + (_s(i) + _s(i - 1)) / 2 * ΔL
                ' If _m(i) < 0 Then _m(i) = 0   'Onnauwkerigheid wordt verdoezeld
            Next

            '=========== Maximaal moment @ imax_count ===============
            Dim temp As Double
            temp = _m(0)
            For i = 0 To _steps
                If _m(i) > temp Then
                    temp = _m(i)
                    imax_count = i
                End If
            Next

            'Debug.WriteLine(_d(imax_count).ToString)

            '=========== Deflection angle, Left hand side ===============
            _α(imax_count) = 0
            For i = imax_count - 1 To 0 Step -1
                _α(i) = _α(i + 1) + _m(i) * ΔL / (2 * Young * pipe_Ix * 10 ^ 6) 'Angle [rad]
                'Debug.WriteLine("Left part i= " & i.ToString & ",  _α(i)= " & _α(i).ToString)
            Next

            '=========== Deflection angle. Right hand side ===============
            _α(imax_count) = 0
            For i = imax_count + 1 To _steps
                _α(i) = _α(i - 1) - _m(i) * ΔL / (2 * Young * pipe_Ix * 10 ^ 6) 'Angle [rad]
                'Debug.WriteLine("Right part i= " & i.ToString & ",  _α(i)= " & _α(i).ToString)
            Next

            '=========== Deflection /sag ==================
            _αv(0) = 0                  'support sag = 0
            For i = 1 To _steps
                _αv(i) = _αv(i - 1) + _α(i) * ΔL * 10 ^ 3 * 2   'Deflection [mm]
            Next

            xnul = _d(imax_count)
            Q_max_bend = _m(imax_count)
            TextBox38.Text = xnul.ToString("F2")          'Positie max moment [m]
            TextBox20.Text = _αv(imax_count).ToString("F1")

            '======= present ==========
            TextBox90.Text = _s(0).ToString("0")             'Shear force
            TextBox4.Text = _s(i_chute_1).ToString("F0")      'Shear force
            TextBox5.Text = _s(i_chute_2).ToString("F0")      'Shear force
            TextBox6.Text = _s(i_chute_3).ToString("F0")      'Shear force
            TextBox91.Text = _s(_steps).ToString("F0")         'Shear force

            TextBox1.Text = _m(i_chute_1).ToString("F0")      'Moment chute #1
            TextBox2.Text = _m(i_chute_2).ToString("F0")      'Moment chute #2
            TextBox3.Text = _m(i_chute_3).ToString("F0")      'Moment chute #3
            TextBox37.Text = Q_max_bend.ToString("F0")  'Max moment [Nm]   

            TextBox114.Clear()
            For i = 0 To _steps    'Write results to text box
                TextBox114.Text &= "Dist=" & _d(i).ToString("F3") & "   "
                TextBox114.Text &= "Shear=" & _s(i).ToString("0000") & "   "
                TextBox114.Text &= "Moment=" & _m(i).ToString("0000") & "   "
                TextBox114.Text &= "Angle=" & (_α(i) * 1000).ToString("F2") & "   "
                TextBox114.Text &= "Sag=" & _αv(i).ToString("F3") & vbCrLf
            Next
            TextBox114.Text &= "Maximum bend moment " & _m(imax_count).ToString("F1") & " [Nm] @ " & xnul.ToString & " [m]" & vbCrLf
            TextBox114.Text &= "Maximum sag " & _αv(imax_count).ToString("F2") & " [mm] "


            '================== calc torsie ========================================
            '=======================================================================
            Tou_torque = P_torque / (pipe_Wp * 1000 ^ 2)        '[N/mm2]
            TextBox12.Text = Tou_torque.ToString("F1")          'Stress from drive [N.m]

            '-------------------------- @ drive max bend------------------------
            P_torque_M = (P_torque * xnul / _λ6)
            Tou_torque_M = P_torque_M / (pipe_Wp * 1000 ^ 2)    '[N/mm2]
            TextBox10.Text = Tou_torque_M.ToString("F1")

            '==================calc stress ===========================================
            '=========================================================================
            '----------- bending stress--------------------
            sigma_eg = Q_max_bend / (pipe_Wx * 1000 ^ 2)        '[N/mm2]
            TextBox09.Text = sigma_eg.ToString("F1")            '[N/mm2]

            '------------ Hubert en hencky @ maximale doorbuiging--------------------
            combined_stress = Sqrt((sigma_eg) ^ 2 + 3 * (Tou_torque_M) ^ 2)
            TextBox21.Text = combined_stress.ToString("F1")

            '---------- allowed sag --------------
            Select Case True
                Case (RadioButton1.Checked)
                    max_sag = 500
                Case (RadioButton2.Checked)
                    max_sag = 800
                Case (RadioButton3.Checked)
                    max_sag = 1000
            End Select


            TextBox49.Text = product_density.ToString("0")

            ' Debug.WriteLine(_αv(imax_count).ToString & " " & (_λ7 * 1000 / max_sag).ToString)
            '---------- checks---------
            TextBox20.BackColor = CType(IIf(_αv(imax_count) > (_λ7 * 1000 / max_sag), Color.Red, Color.LightGreen), Color)
            TextBox09.BackColor = CType(IIf(sigma_eg > sigma_fatique, Color.Red, Color.LightGreen), Color)
            TextBox21.BackColor = CType(IIf(combined_stress > sigma_fatique, Color.Red, Color.LightGreen), Color)
            TextBox12.BackColor = CType(IIf(Tou_torque > sigma_fatique, Color.Red, Color.LightGreen), Color)
            NumericUpDown28.BackColor = CType(IIf(_λ4 > _λ6, Color.Red, Color.Yellow), Color) 'Inlet #3
            NumericUpDown24.BackColor = CType(IIf(_λ3 > _λ4, Color.Red, Color.Yellow), Color) 'Inlet #2
            NumericUpDown16.BackColor = CType(IIf(_λ2 > _λ3, Color.Red, Color.Yellow), Color) 'Inlet #1
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, NumericUpDown11.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown13.ValueChanged, TabPage2.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, ComboBox5.SelectedIndexChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, CheckBox1.CheckedChanged, NumericUpDown18.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown28.ValueChanged, CheckBox10.CheckedChanged, NumericUpDown29.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown22.ValueChanged
        Calc_sequence()
    End Sub
    Private Sub Screw_dia_combo()
        Dim words() As String

        If (ComboBox11.SelectedIndex > -1) Then      'Prevent exceptions
            words = Flight_dia(ComboBox11.SelectedIndex).Split(CType(";", Char()))
            Double.TryParse(words(0), _diam_flight)
            _diam_flight /= 1000                    'Trough width[m]

            TextBox18.Text = (_diam_flight * 1000).ToString("F3")
            MessageBox.Show(_diam_flight.ToString)
        End If
    End Sub

    Private Sub Pipe_dia_combo_init()
        Dim words(), tmp As String

        ComboBox3.Items.Clear()
        ComboBox9.Items.Clear()

        '-------Fill combobox3, Pipe selection------------------
        For hh = 0 To (UBound(pipe_steel) - 1)                'Fill combobox 3 with pipe data
            words = pipe_steel(hh).Split(CType(";", Char()))
            ComboBox3.Items.Add(Trim(words(2)))
            ComboBox9.Items.Add(Trim(words(2)))
            tmp = "     " & Trim(words(2))
            ComboBox14.Items.Add(tmp)
        Next hh
        ComboBox3.SelectedIndex = 5
        ComboBox9.SelectedIndex = 2
        ComboBox14.SelectedIndex = 2

        words = pipe_steel(ComboBox3.SelectedIndex).Split(CType(";", Char()))
        Double.TryParse(words(2), _pipe_OD)
        _pipe_OD /= 1000                                         'Outside Diameter [mm]
        TextBox16.Text = (_pipe_OD * 1000).ToString("F1")     'Diameter [mm]
    End Sub
    Private Sub Pipe_wall_combo_init()
        Dim words() As String
        Dim temp As Double

        If ComboBox3.Items.Count > 0 Then 'Prevent exceptions
            ComboBox6.Items.Clear()
            '-------Fill combobox6, pipe wall selection------------------
            words = pipe_steel(ComboBox3.SelectedIndex).Split(CType(";", Char()))  'Fill combobox 6 pipe wall data
            For hh = 3 To 6
                Double.TryParse(words(hh), temp)
                If temp > 0 Then ComboBox6.Items.Add(Trim(words(hh)))
            Next
            ComboBox6.SelectedIndex = 0
            Calulate_stress_1()
        End If
    End Sub
    Private Sub Motorreductor()
        Dim words() As String

        ComboBox4.Items.Clear()
        '-------Fill combobox4,  selection------------------
        For hh = 1 To (UBound(motorred))                'Fill combobox 3 with pipe data
            words = motorred(hh).Split(CType(";", Char()))
            ComboBox4.Items.Add(Trim(words(0)))
        Next hh
        ComboBox4.SelectedIndex = 2
    End Sub
    Private Sub Coupling_combo()
        Dim words() As String

        ComboBox7.Items.Clear()
        '-------Fill combobox7,  selection------------------
        For hh = 1 To (UBound(coupl))                'Fill combobox 7 with coupling data
            words = coupl(hh).Split(CType(";", Char()))
            ComboBox7.Items.Add(Trim(words(0)))
        Next hh
        ComboBox7.SelectedIndex = 1
    End Sub
    Private Sub Lager_combo()
        Dim words() As String

        ComboBox8.Items.Clear()
        '-------Fill combobox8,  selection------------------
        For hh = 1 To lager.Length - 1                'Fill combobox 8 with lager data
            words = lager(hh).Split(CType(";", Char()))
            ComboBox8.Items.Add(words(0))
        Next hh
        ComboBox8.SelectedIndex = 1
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row, font_sizze As Integer

        'Start Word and open the document template. 
        font_sizze = 9
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 3
        oPara1.Range.Font.Bold = CInt(True)
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze + 1
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = CInt(False)
        oPara2.Range.Text = "Screw conveyor cost calculation" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox23.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox35.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Conveyor id"
        oTable.Cell(row, 2).Range.Text = TextBox53.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 14, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input Data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter Flight"
        oTable.Cell(row, 3).Range.Text = ComboBox11.Text
        oTable.Cell(row, 2).Range.Text = "[mm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter pipe"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox9.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox45.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Wall thickness pipe"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox6.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Pitch"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown2.Value, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Flight thickness"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox46.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Conveyor speed"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown7.Value, String)
        oTable.Cell(row, 2).Range.Text = "[rpm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Installed Power"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox5.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[kW]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Service factor"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown18.Value, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Conveyor length"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown3.Value, String)
        oTable.Cell(row, 2).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Inclination angle"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown4.Value, String)
        oTable.Cell(row, 2).Range.Text = "[deg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Steel type"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox2.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Temperature"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown11.Value, String)
        oTable.Cell(row, 2).Range.Text = "[C]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Product type"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox1.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        '---- -----
        oTable.Cell(row, 1).Range.Text = "Capacity"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown5.Value, String)
        oTable.Cell(row, 2).Range.Text = "[ton/hr]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(1.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.6)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 12, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Gearreducer"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox4.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Coupling"
        NewMethod(oTable, row)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Bearing shaft diameter"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox8.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "No. certificates "
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown27.Value, String)
        oTable.Cell(row, 2).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Weight end plates"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown10.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox42.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Weight flighting"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox46.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Weight trough"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown14.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox47.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Weight cover"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown15.Value, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox48.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Weight stub shafts"
        oTable.Cell(row, 3).Range.Text = CType(ComboBox13.SelectedItem, String)
        oTable.Cell(row, 2).Range.Text = "[mm]"
        oTable.Cell(row, 5).Range.Text = TextBox54.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        row += 1
        oTable.Cell(row, 5).Range.Text = "_____"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Total sheet steel"
        oTable.Cell(row, 5).Range.Text = TextBox109.Text
        oTable.Cell(row, 4).Range.Text = "[kg]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(1.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(0.6)

        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 14, 8)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Costs"
        row += 1
        oTable.Rows.Item(2).Range.Font.Bold = CInt(True)
        oTable.Cell(row, 6).Range.Text = "Material"
        oTable.Cell(row, 1).Range.Text = "Labour"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Work preparation"
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown48.Value, String)
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 5).Range.Text = TextBox140.Text

        row += 1
        oTable.Cell(row, 1).Range.Text = "Engineering"
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown30.Value, String)
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 5).Range.Text = TextBox55.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project man."
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown33.Value, String)
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 5).Range.Text = TextBox70.Text

        row += 1
        oTable.Cell(row, 1).Range.Text = "Work shop"
        oTable.Cell(row, 3).Range.Text = CType(NumericUpDown34.Value, String)
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 5).Range.Text = TextBox72.Text
        oTable.Cell(row, 4).Range.Text = "[€]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Total hours"
        oTable.Cell(row, 3).Range.Text = TextBox106.Text
        oTable.Cell(row, 2).Range.Text = "[hr]"
        oTable.Cell(row, 5).Range.Text = TextBox98.Text
        oTable.Cell(row, 4).Range.Text = "[€]"
        oTable.Cell(row, 6).Range.Text = "Total cost material"
        oTable.Cell(row, 8).Range.Text = TextBox103.Text
        oTable.Cell(row, 7).Range.Text = "[€]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Percentage labour "
        oTable.Cell(row, 5).Range.Text = TextBox101.Text
        oTable.Cell(row, 4).Range.Text = "[%]"
        oTable.Cell(row, 6).Range.Text = "Percentage material"
        oTable.Cell(row, 8).Range.Text = TextBox100.Text
        oTable.Cell(row, 7).Range.Text = "[%]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Total cost price"
        oTable.Cell(row, 5).Range.Text = TextBox73.Text
        oTable.Cell(row, 4).Range.Text = "[€]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Packing"
        oTable.Cell(row, 5).Range.Text = CType(NumericUpDown49.Value, String)
        oTable.Cell(row, 4).Range.Text = "[€]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Shipping"
        oTable.Cell(row, 5).Range.Text = CType(NumericUpDown50.Value, String)
        oTable.Cell(row, 4).Range.Text = "[€]"

        row += 1
        oTable.Rows.Item(12).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(12).Range.Font.Size = font_sizze + 1
        oTable.Cell(row, 1).Range.Text = "Total sale price"
        oTable.Cell(row, 5).Range.Text = TextBox75.Text
        oTable.Cell(row, 4).Range.Text = "[€]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(1.2)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.6)
        oTable.Columns.Item(4).Width = oWord.InchesToPoints(0.4)   'Change width of columns 1 & 2.
        oTable.Columns.Item(5).Width = oWord.InchesToPoints(1.4)
        oTable.Columns.Item(6).Width = oWord.InchesToPoints(1.5)
        oTable.Columns.Item(7).Width = oWord.InchesToPoints(0.4)
        oTable.Columns.Item(8).Width = oWord.InchesToPoints(0.6)
    End Sub

    Private Sub NewMethod(oTable As Word.Table, row As Integer)
        oTable.Cell(row, 3).Range.Text = CType(ComboBox7.SelectedItem, String)
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs)
        Screw_dia_combo()
        Calculate()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox65.Text.Trim.Length > 0 And TextBox66.Text.Trim.Length > 0 And TextBox66.Text.Trim.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Complete Project number, name and item number")
        End If
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string, user As String

        user = Trim(Environment.UserName)         'User name on the screen
        Dim filename As String
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        filename = "Conveyor_select_" & TextBox66.Text & "_" & TextBox65.Text & "_" & TextBox67.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtks"

        temp_string = TextBox66.Text & ";" & TextBox65.Text & ";" & TextBox67.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(System.Windows.Forms.ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As System.Windows.Forms.ComboBox = CType(all_combo(i), System.Windows.Forms.ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        Check_directories()  'Are the directories present
        If CInt(temp_string.Length.ToString) > 5 Then      'String may be empty
            If Directory.Exists(dirpath_Eng) Then
                File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
            Else
                File.WriteAllText(dirpath_Home_GP & filename, temp_string, Encoding.ASCII)     'used at home
            End If
        End If
    End Sub
    Private Sub Check_directories()
        '---- if path not exist then create one----------
        If (Not System.IO.Directory.Exists(dirpath_Home_GP)) Then System.IO.Directory.CreateDirectory(dirpath_Home_GP)
        If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
        If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)

    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Read_file()
        Calculate()
    End Sub
    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim k As Integer = 0
        Dim ttt As Double
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Conveyor_select_"

        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home_GP  'used at home
        End If

        OpenFileDialog1.Title = "Open a VTKS File"
        OpenFileDialog1.Filter = "VTKQ Files|*.vtks|VTKQ file|*.vtks"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)

            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split the read file content
            TextBox66.Text = words(0)                  'Project number
            TextBox65.Text = words(1)                  'Project name
            TextBox67.Text = words(2)                  'Item no

            '---------- terugzetten Numeric controls -----------------
            FindControlRecursive(all_num, Me, GetType(System.Windows.Forms.NumericUpDown))
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As System.Windows.Forms.NumericUpDown = CType(all_num(i), System.Windows.Forms.NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        grbx.Value = grbx.Minimum       'NOK
                        MessageBox.Show("Numeric controls value out of min-max range, Minimum value is used")
                    End If
                Else
                    MessageBox.Show("Warning last Numeric-Updown-controls not found in file")  'NOK
                End If
            Next

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(System.Windows.Forms.ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()          'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As System.Windows.Forms.ComboBox = CType(all_combo(i), System.Windows.Forms.ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(System.Windows.Forms.RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As System.Windows.Forms.RadioButton = CType(all_radio(i), System.Windows.Forms.RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last radiobutton not found in file")
                End If
            Next

        End If
    End Sub
    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        Pipe_wall_combo_init()
        Calulate_stress_1()
    End Sub

    Private Sub ComboBox6_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedValueChanged
        Calulate_stress_1()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles TabPage6.Enter, Button3.Click
        Draw_chart1()
        Draw_chart2()
        Draw_chart3()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, TabPage7.Enter, NumericUpDown38.ValueChanged, NumericUpDown26.ValueChanged
        Dim height, weight, speed, time, force, acc As Double

        weight = NumericUpDown38.Value
        height = NumericUpDown26.Value

        time = Sqrt(2 * height / 9.81)
        speed = Sqrt(2 * 9.81 * height)
        acc = speed / 0.01
        force = weight * acc


        TextBox119.Text = time.ToString("0.0")
        TextBox120.Text = speed.ToString("0.0")
        TextBox121.Text = acc.ToString("0.0")
        TextBox122.Text = force.ToString("0")
    End Sub

    Private Sub Astap_combo()
        Dim words() As String

        ComboBox13.Items.Clear()
        '-------Fill combobox------------------
        For hh = 1 To astap_dia.Length - 1                'Fill combobox 13 with astap data
            words = astap_dia(hh).Split(CType(";", Char()))
            ComboBox13.Items.Add(words(0))
        Next hh
        ComboBox13.SelectedIndex = 1
    End Sub

    Private Sub Screw_combo_init()
        Dim words() As String

        ComboBox11.Items.Clear()
        '-------Fill combobox------------------
        For hh = 1 To Flight_dia.Length - 1                'Fill combobox 11 with flight data
            words = Flight_dia(hh).Split(CType(";", Char()))
            ComboBox11.Items.Add(words(0))
        Next hh
        ComboBox11.SelectedIndex = 2
    End Sub
    Private Sub Paint_combo()
        Dim words() As String

        ComboBox12.Items.Clear()
        '-------Fill combobox ------------------
        For hh = 1 To ppaint.Length - 1         'Fill combobox 12 with paint data
            words = ppaint(hh).Split(CType(";", Char()))
            ComboBox12.Items.Add(words(0))
        Next hh
        ComboBox12.SelectedIndex = 0            'Pickling + passivating
    End Sub
    Private Sub Pakking_combo()
        Dim words() As String

        ComboBox10.Items.Clear()
        '-------Fill combobox-----------------
        For hh = 1 To pakking.Length - 1                'Fill combobox 10 with pakking data
            words = pakking(hh).Split(CType(";", Char()))
            ComboBox10.Items.Add(words(0))
        Next hh
        ComboBox10.SelectedIndex = 3
    End Sub
    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        Calculate()
        Calulate_stress_1()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown41.ValueChanged, TabPage10.Enter, NumericUpDown45.ValueChanged, ComboBox14.SelectedIndexChanged
        'Calculate flight weight, everything in [m]
        Dim blank_Dia As Double     '[mm]
        Dim blank_wgt As Double     '[kg]
        Dim blank_cost As Double    '[euro] material cost

        conv.flight_OD = NumericUpDown43.Value / 1000           '[m] OD flight
        Double.TryParse(ComboBox14.Text, conv.pipe_id)          '[mm] OD pipe
        conv.pipe_id /= 1000                                    '[m] OD pipe

        conv.pitch = NumericUpDown42.Value * conv.pipe_od       '[m] pitch
        conv.flight_thick = NumericUpDown41.Value / 1000        '[m] flight thickness


        Debug.WriteLine("conv.pipe_od= " & conv.pipe_od.ToString)

        '---------- blank dimensions before forming ----
        blank_Dia = Blank_OD(conv.flight_OD, conv.pitch)          '[m] 

        Debug.WriteLine(" blank_Dia= " & blank_Dia.ToString)

        blank_wgt = blank_Dia ^ 2 * conv.flight_thick * 7850    '[kg]
        blank_cost = blank_wgt * NumericUpDown45.Value          '[e]

        '---------- weight of one 360 degree flight -----
        Flight_weight(conv, 1)   'calc flight weight

        '---------- present results ----------
        TextBox128.Text = conv.flight_weight.ToString("F2")     '[kg]
        TextBox130.Text = (blank_Dia * 1000).ToString("F0")     '[mm]
        TextBox132.Text = blank_wgt.ToString("F2")              '[kg]
        TextBox131.Text = blank_cost.ToString("F2")             '[e]
    End Sub
    ' Private Function Flight_weight(d1 As Double, d2 As Double, pitch As Double, thick As Double, no_f As Double) As Double
    Private Sub Flight_weight(ByRef c As Conveyor_struct, no_f As Double)
        Dim hoek_spoed As Double

        Dim tip_length As Double

        hoek_spoed = Atan(c.pitch / (PI * c.flight_OD))            '[rad]  
        tip_length = Sqrt(c.pitch ^ 2 + c.flight_OD ^ 2)           '[m]

        c.flight_weight = PI / 4 * 7850 * c.flight_thick * no_f * (c.flight_OD ^ 2 - c.pipe_od ^ 2) / Cos(hoek_spoed)

        'Debug.WriteLine("============== ")
        'Debug.WriteLine("pitch= " & c.pitch.ToString)
        'Debug.WriteLine("c.dia_flight= " & c.flight_OD.ToString)
        'Debug.WriteLine("c.pipe_od= " & c.pipe_od.ToString)
        'Debug.WriteLine("c.flight_weight= " & c.flight_weight.ToString)
    End Sub
    Private Function Blank_OD(d1 As Double, pitch As Double) As Double
        'Weight of the square blank
        Dim blank_dia As Double     'diameter flight blank (before forming)
        Dim tip_length As Double    'outside length 1 flight

        tip_length = Sqrt(pitch ^ 2 + (PI * d1) ^ 2)   '[mm]

        blank_dia = tip_length / PI             '[mm]
        Return (blank_dia)
    End Function

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click, TabPage11.Enter, NumericUpDown47.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown12.ValueChanged
        Dim dia As Double           '[mm] screw conveyor
        Dim length As Double        '[m] screw
        Dim speed As Double         '[rpm] screw
        Dim T_gearbox As Double     '[Nm] Torque gearbox
        Dim power_motor As Double   '[kW]
        Dim eff As Double = 0.8     '[-]

        '----- Get data -------
        dia = NumericUpDown44.Value         '[mm] screw conveyor
        length = NumericUpDown12.Value      '[m] screw
        speed = NumericUpDown47.Value       '[rpm] screw

        '----- Calc torque ---------
        T_gearbox = Calc_NON_Torque(dia, length)

        '----- Calc power ----------
        power_motor = (T_gearbox * speed * 2 * PI / 60) / eff   '[W]
        power_motor /= 1000                                 '[kW]

        TextBox141.Text = T_gearbox.ToString("F0")          '[Nm]
        TextBox142.Text = power_motor.ToString("F1")        '[kW]
    End Sub
    Private Function Calc_NON_Torque(dia As Double, length As Double) As Double
        'Based on chart from Noord-Oost Nederland, Asperen
        Dim rc As Double            '[richtingscoeficient]
        Dim offset As Double        '[verticale nul verschuiving]
        Dim torque As Double

        rc = 0.00029 * dia ^ 2 - 0.17518 * dia + 74.562     '[-]
        offset = 0.8229 * dia - 200.25                      '[-]
        torque = rc * length + offset                       '[Nm]
        Return (torque)
    End Function
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Print_word()
    End Sub

    Private Sub Print_word()
        Dim oWord As Word.Application ' = Nothing
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, opara3 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename, str As String
        Dim speed As Double

        oWord = New Word.Application()

        'Start Word and open the document template. 
        font_sizze = 8
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True
        oDoc = oWord.Documents.Add
        oDoc.PageSetup.TopMargin = 35
        oDoc.PageSetup.BottomMargin = 20
        oDoc.PageSetup.RightMargin = 20
        oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait
        oDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 3
        oPara1.Range.Font.Bold = CInt(True)
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze + 1
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = CInt(False)
        oPara2.Range.Text = "Screw Conveyor Stress calculation " & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 0
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox66.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox65.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Machine Id "
        oTable.Cell(row, 2).Range.Text = TextBox67.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns(2).Width = oWord.InchesToPoints(2)

        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 0
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Material"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Product"
        str = ComboBox1.Text
        If Len(str) > 22 Then str = str.Substring(0, 22)
        oTable.Cell(row, 2).Range.Text = str

        row += 1
        oTable.Cell(row, 1).Range.Text = "Product Flow"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown5.Value, String)
        oTable.Cell(row, 3).Range.Text = "[ton/hr]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Product Density"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown6.Value, String)
        oTable.Cell(row, 3).Range.Text = "[kg/m3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Forward resistance"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown9.Value, String)
        oTable.Cell(row, 3).Range.Text = "[-]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
        oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
        oTable.Columns(3).Width = oWord.InchesToPoints(1.5)

        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 13, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Conveyor Data"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter flight"
        oTable.Cell(row, 2).Range.Text = ComboBox11.Text
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter pipe"
        oTable.Cell(row, 2).Range.Text = ComboBox3.Text & " x " & ComboBox6.Text
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Pitch"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown2.Value, String)
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Blade thickness"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Length"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown3.Value, String)
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Inclination angle"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown4.Value, String)
        oTable.Cell(row, 3).Range.Text = "[degree]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Speed"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown7.Value, String)
        oTable.Cell(row, 3).Range.Text = "[rpm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Flight tip speed"
        oTable.Cell(row, 2).Range.Text = TextBox11.Text
        oTable.Cell(row, 3).Range.Text = "[m/s]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Filling "
        oTable.Cell(row, 2).Range.Text = TextBox01.Text
        oTable.Cell(row, 3).Range.Text = "[%]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Power ISO 7119"
        oTable.Cell(row, 2).Range.Text = TextBox03.Text
        oTable.Cell(row, 3).Range.Text = "[kW]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Power MEKOG"
        oTable.Cell(row, 2).Range.Text = TextBox04.Text
        oTable.Cell(row, 3).Range.Text = "[kW]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Power Installed"
        oTable.Cell(row, 2).Range.Text = ComboBox5.Text
        oTable.Cell(row, 3).Range.Text = "[kW]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
        oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
        oTable.Columns(3).Width = oWord.InchesToPoints(1.5)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '-------------Dimensions inlets-------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Inlets chutes"
        oTable.Cell(row, 2).Range.Text = ""
        oTable.Cell(row, 3).Range.Text = ""
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #1 size and location"
        oTable.Cell(row, 2).Range.Text = ComboBox11.Text & " x " & (NumericUpDown31.Value * 1000).ToString("0") & " @ " & (NumericUpDown16.Value * 1000).ToString
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #2 size and location"
        oTable.Cell(row, 2).Range.Text = ComboBox11.Text & " x " & (NumericUpDown32.Value * 1000).ToString("0") & " @ " & (NumericUpDown24.Value * 1000).ToString
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #3 size and location"
        oTable.Cell(row, 2).Range.Text = ComboBox11.Text & " x " & (NumericUpDown22.Value * 1000).ToString("0") & " @ " & (NumericUpDown28.Value * 1000).ToString
        oTable.Cell(row, 3).Range.Text = "[mm]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
        oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
        oTable.Columns(3).Width = oWord.InchesToPoints(1.5)

        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        '-------------Loads-------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Chute loads"
        oTable.Cell(row, 2).Range.Text = ""
        oTable.Cell(row, 3).Range.Text = ""

        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #1 material column"
        oTable.Cell(row, 2).Range.Text = NumericUpDown19.Value.ToString
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #1 Force"
        oTable.Cell(row, 2).Range.Text = TextBox115.Text
        oTable.Cell(row, 3).Range.Text = "[N]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #2 material column"
        oTable.Cell(row, 2).Range.Text = NumericUpDown36.Value.ToString
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #2 Force"
        oTable.Cell(row, 2).Range.Text = TextBox116.Text
        oTable.Cell(row, 3).Range.Text = "[N]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #3 material column"
        oTable.Cell(row, 2).Range.Text = NumericUpDown37.Value.ToString
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Chute #3 Force"
        oTable.Cell(row, 2).Range.Text = TextBox117.Text
        oTable.Cell(row, 3).Range.Text = "[N]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Uniform material load"
        oTable.Cell(row, 2).Range.Text = TextBox118.Text
        oTable.Cell(row, 3).Range.Text = "[N/m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Pipe and flight weight only"
        oTable.Cell(row, 2).Range.Text = TextBox22.Text
        oTable.Cell(row, 3).Range.Text = "[N/m]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
        oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
        oTable.Columns(3).Width = oWord.InchesToPoints(1.5)

        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------- Results----------------------
        'Insert a 5 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Calculation Results"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Bending Stress"
        oTable.Cell(row, 2).Range.Text = TextBox09.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Max. Torque Stress"
        oTable.Cell(row, 2).Range.Text = TextBox12.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Combined Stress"
        oTable.Cell(row, 2).Range.Text = TextBox21.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Selected steel"
        oTable.Cell(row, 2).Range.Text = ComboBox2.Text
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Max. allowed Fatique stress"
        oTable.Cell(row, 2).Range.Text = TextBox08.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Maximum Flex"
        oTable.Cell(row, 2).Range.Text = TextBox20.Text
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Maximum Flex @"
        oTable.Cell(row, 2).Range.Text = TextBox38.Text
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)
        oTable.Columns(2).Width = oWord.InchesToPoints(1.8)
        oTable.Columns(3).Width = oWord.InchesToPoints(1.5)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------------save picture #1---------------- 
        Draw_chart1()
        Chart1.SaveImage(dirpath_Home_GP & "ShearChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
        opara3 = oDoc.Content.Paragraphs.Add
        opara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify
        opara3.Range.InlineShapes.AddPicture(dirpath_Home_GP & "ShearChart.gif")
        opara3.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
        opara3.Range.InlineShapes.Item(1).ScaleWidth = 30       'Size
        opara3.Range.InsertParagraphAfter()

        '------------------save picture #2 ---------------- 
        Draw_chart2()
        Chart2.SaveImage(dirpath_Home_GP & "DeflectionradChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
        opara3 = oDoc.Content.Paragraphs.Add
        opara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify
        opara3.Range.InlineShapes.AddPicture(dirpath_Home_GP & "DeflectionradChart.gif")
        opara3.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
        opara3.Range.InlineShapes.Item(1).ScaleWidth = 30       'Size
        opara3.Range.InsertParagraphAfter()

        '------------------save picture #3 ---------------- 
        Draw_chart3()
        Chart3.SaveImage(dirpath_Home_GP & "DeflectionmmChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
        opara3 = oDoc.Content.Paragraphs.Add
        opara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify
        opara3.Range.InlineShapes.AddPicture(dirpath_Home_GP & "DeflectionmmChart.gif")
        opara3.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
        opara3.Range.InlineShapes.Item(1).ScaleWidth = 30       'Size
        opara3.Range.InsertParagraphAfter()

        ''-------------- Checks-------
        'Insert a 5 x 1 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 1)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
        row = 1
        oTable.Cell(row, 1).Range.Text = "Checks "
        row += 1

        Double.TryParse(TextBox11.Text, speed)
        If (speed > 1.0) Then
            oTable.Cell(row, 1).Range.Text = "NOK, for ATEX, Flight speed > 1 m/s"
        Else
            oTable.Cell(row, 1).Range.Text = "OK, Flight speed for ATEX applications"
        End If
        row += 1
        If NumericUpDown7.BackColor = Color.Red Then
            oTable.Cell(row, 1).Range.Text = "NOK, Rotational speed > 45 rpm, too fast"
        Else
            oTable.Cell(row, 1).Range.Text = "OK, Rotational speed"
        End If
        row += 1
        If TextBox01.BackColor = Color.Red Then
            oTable.Cell(row, 1).Range.Text = "NOK, Filling percentage > 45%"
        Else
            oTable.Cell(row, 1).Range.Text = "OK, Filling percentage"
        End If
        row += 1
        If TextBox21.BackColor = Color.Red Then
            oTable.Cell(row, 1).Range.Text = "NOK, Combined pipe stress too high"
        Else
            oTable.Cell(row, 1).Range.Text = "OK, Combined pipe stress"
        End If
        row += 1
        If TextBox20.BackColor = Color.Red Then
            oTable.Cell(row, 1).Range.Text = "NOK, Deflection pipe too high"
        Else
            oTable.Cell(row, 1).Range.Text = "OK, Deflection pipe stress"
        End If
        oTable.Columns(1).Width = oWord.InchesToPoints(4.0)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '--------------Save file word file------------------
        'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx

        ufilename = "Conveyor_report_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"

        If Directory.Exists(dirpath_Rap) Then
            ufilename = dirpath_Rap & ufilename
        Else
            ufilename = dirpath_Home_GP & ufilename
        End If
        oWord.ActiveDocument.SaveAs(ufilename.ToString)


    End Sub

    Private Sub Costing_material()
        Dim rho_staal, rho_kunststof As Double
        Dim opp_trog, opp_kopstaartplaat, kg_kopstaartplaat, kg_trog As Double
        Dim kg_pipe, speling_trog, diam_schroef As Double
        Dim kopstaart_dikte, dikte_deksel, dikte_lining, dikte_trog As Double
        Dim kg_inlaat, kg_uitlaat, kg_lining As Double
        Dim kg_deksel, kg_voet, kg_schroefblad, total_kg_plaat As Double
        Dim hoek_spoed As Double
        Dim nr_flights, spoed As Double
        Dim kg_astap, dia_astap, lengte_astap As Double
        Dim kg_afschermkap As Double
        Dim tot_opperv_paint As Double
        Dim oppb_afschermkap, oppb_astap, oppb_voet, oppb_uitlaat, oppb_inlaat As Double
        Dim oppb_deksel, oppb_trog, oppb_kopstaartplaat, oppb_schroefblad, oppb_pipe As Double
        Dim cost_kopstaartplaat, cost_trog, cost_pipe, cost_deksel, cost_inlaat, cost_uitlaat As Double
        Dim cost_voet, cost_schroefblad, cost_astap, cost_lining, cost_afschermkap As Double
        Dim cost_paint, cost_painting, cost_cutting As Double
        Dim cost_motorreductor, cost_koppeling, cost_lagers As Double
        Dim cost_pakking, cost_hang, cost_transport As Double
        Dim cost_stopbus As Double
        Dim certificate_cost, total_cost As Double
        Dim uren_wvb, uren_eng, uren_pro, uren_fab, tot_uren As Double
        Dim eng_prijs_uur, project_prijs_uur, fabriek_prijs_uur, wvb_prijs_uur As Double
        Dim prijs_wvb, prijs_eng, prijs_pro, prijs_fab As Double
        Dim tot_prijsarbeid, geheel_totprijs, dekking, marge_cost, verkoopprijs As Double
        Dim perc_mater, perc_arbeid As Double
        Dim gew_inuitvoet As Double
        Dim packing, shipping As Double                 'packing, shipping costs

        TextBox40.Text = ComboBox2.Text                'materiaalsoort staal
        TextBox41.Text = (_pipe_OD * 1000).ToString    'diameter pijp
        TextBox51.Text = CType(NumericUpDown3.Value, String)          'lengte trog
        TextBox52.Text = ComboBox5.Text                'vermogen aandrijving
        TextBox44.Text = _diam_flight.ToString         'diameter flight

        '---------------------------------------------- PRICES -----------------------------------------
        '-----------------------------------------------------------------------------------------------
        TextBox84.Text = "3.25"                 'lining [€/kg]
        TextBox85.Text = "10.00"                'alu afschermkap [€/kg]

        Select Case True
            Case (RadioButton6.Checked)         'staal, s235JR
                rho_staal = 7850
                TextBox93.Text = "0.78"         'kop staart  [€/kg]
                TextBox96.Text = "0.78"         'trog
                TextBox97.Text = "0.78"         'deksel
                TextBox113.Text = "0.78"        'inlaat,uitlaat,voet, schermkap [€/kg]
                TextBox94.Text = CType(IIf(Not CheckBox5.Checked, "2.00", "3.00"), String) 'schroefpijp staal (seam/seamless)
                TextBox95.Text = "8.00"         'schroefblad
                TextBox92.Text = "2.09"         'astap ronde staf afm 60
                CheckBox6.Checked = True        'Paint
            Case (RadioButton7.Checked)         'rvs304, (Koud + 2B)
                rho_staal = 7950
                TextBox93.Text = "2.75"         'kop staart 
                TextBox96.Text = "2.75"         'trog
                TextBox97.Text = "2.75"         'deksel
                TextBox113.Text = "2.75"        'inlaat,uitlaat,voet, schermkap [€/kg]
                TextBox94.Text = CType(IIf(Not CheckBox5.Checked, "3.45", "6.25"), String) 'schroefpijp 304 (seam/seamless)
                TextBox95.Text = "10.0"         'schroefblad
                TextBox92.Text = "1.52"         'astap [€/kg] materiaal is standaard van staal
                CheckBox6.Checked = False       'Paint
            Case (RadioButton8.Checked)         'rvs316, (Koud + 2B)
                rho_staal = 8000
                TextBox93.Text = "3.97"         'kop staart 
                TextBox96.Text = "3.97"         'trog
                TextBox97.Text = "3.97"         'deksel
                TextBox113.Text = "3.97"        'inlaat,uitlaat,voet, schermkap [€/kg]
                TextBox94.Text = CType(IIf(Not CheckBox5.Checked, "4.45", "8.20"), String) 'schroefpijp 316 (seam/seamless)
                TextBox95.Text = "10.00"        'schroefblad
                TextBox92.Text = "1.52"         'astap [€/kg] materiaal is standaard van staal
                CheckBox6.Checked = False       'Paint
        End Select

        If ComboBox4.SelectedIndex > -1 And ComboBox7.SelectedIndex > -1 And ComboBox8.SelectedIndex > -1 And
            ComboBox10.SelectedIndex > -1 And ComboBox12.SelectedIndex > -1 Then
            Dim words3() As String = motorred(ComboBox4.SelectedIndex + 1).Split(CType(";", Char()))
            Dim words2() As String = coupl(ComboBox7.SelectedIndex + 1).Split(CType(";", Char()))
            Dim words1() As String = lager(ComboBox8.SelectedIndex + 1).Split(CType(";", Char()))
            Dim words5() As String = pakking(ComboBox10.SelectedIndex + 1).Split(CType(";", Char()))
            Dim words4() As String = ppaint(ComboBox12.SelectedIndex + 1).Split(CType(";", Char()))

            cost_lagers = CDbl(words1(1))
            cost_koppeling = CDbl(words2(1)) * CDbl(words2(2))       'inclusief kortingspercentage van 45%
            If Not CheckBox3.Checked Then cost_koppeling = 0
            cost_motorreductor = CDbl(words3(3))
            If Not CheckBox2.Checked Then cost_motorreductor = 0
            cost_paint = CDbl(words4(1))
            If Not CheckBox6.Checked Then cost_paint = 0
            cost_pakking = CDbl(words5(1))
            cost_inlaat = 300   'inlaat chute
            cost_uitlaat = 300  'Uitlaat chute
            cost_voet = 100     'Conveyor supports
        End If


        '---------------Plaat diktes---------------
        dikte_trog = NumericUpDown14.Value / 1000
        kopstaart_dikte = NumericUpDown10.Value / 1000  '[m]
        If RadioButton5.Checked Then
            dikte_deksel = NumericUpDown15.Value / 1000 'U-trog
        Else
            dikte_deksel = 0                            'Pijpschroef
        End If
        dikte_lining = NumericUpDown25.Value / 1000

        '--------------staal Oppervlaktes -------
        opp_trog = PI * _diam_flight * _λ6
        opp_kopstaartplaat = _diam_flight ^ 2

        '--------------paint Oppervlaktes -------
        oppb_kopstaartplaat = 2 * opp_kopstaartplaat               'Binnen en buiten
        oppb_trog = 2 * (opp_kopstaartplaat + opp_trog)      'kuip zowel uitwendig als inwendig
        oppb_pipe = _pipe_OD / 1000 * PI * _λ6

        '-------------- plaat gewichten---------------
        kg_kopstaartplaat = _diam_flight ^ 2 * kopstaart_dikte * rho_staal
        kg_trog = _diam_flight * 4 * dikte_trog * _λ6 * rho_staal
        kg_deksel = _diam_flight * 1.1 * dikte_deksel * _λ6 * rho_staal     '50mm voor de horizontale flens en 25mm voor het stukje naar beneden

        'Debug.WriteLine(" kg_kopstaartplaat= " & kg_kopstaartplaat.ToString)
        'Debug.WriteLine(" kg_trog= " & kg_trog.ToString)
        'Debug.WriteLine(" kg_deksel= " & kg_deksel.ToString)

        '--------------- pipe gewicht-------------------
        Double.TryParse(CType(ComboBox9.SelectedItem, String), _pipe_OD)         ' ComboBox3 = ComboBox9
        _pipe_OD /= 1000                  '[m]
        _pipe_wall = CDbl(ComboBox6.SelectedItem)
        _pipe_wall /= 1000                          '[m]
        _pipe_ID = (_pipe_OD - 2 * _pipe_wall)
        kg_pipe = rho_staal * PI / 4 * (_pipe_OD ^ 2 - _pipe_ID ^ 2) * _λ6

        '----------------------------------------WEIGHT + AREA CALCULATIONS-----------------------------------------
        '----------------------------------------------------------------------------------------------
        If _diam_flight > 0.3015 Then   'in [m], radiale speling schroef in kuip: tot diam 0.3m 7.5 mm, daarboven 10mm
            speling_trog = 0.01
        Else
            speling_trog = 0.0075
        End If
        diam_schroef = _diam_flight - 2 * speling_trog
        oppb_deksel = 2 * _λ6 * (_diam_flight + 0.075)                  'zowel inwendig als uitwendig

        ' NumericUpDown12.Value = NumericUpDown8.Value                    'Dikte schroefblad bij tab1 opgegeven
        spoed = diam_schroef * NumericUpDown2.Value
        nr_flights = _λ6 / spoed
        hoek_spoed = Atan(spoed / (PI * diam_schroef))                  '[rad]    

        kg_schroefblad = PI * rho_staal * (NumericUpDown8.Value / 1000) * 0.25 * nr_flights * (diam_schroef ^ 2 - _pipe_OD ^ 2) / Cos(hoek_spoed)         ' DIT IS DE ECHTE FORMULE!!!!!
        oppb_schroefblad = 2 * (kg_schroefblad / (NumericUpDown8.Value * rho_staal / 1000))

        Double.TryParse(CType(ComboBox13.SelectedItem, String), dia_astap)             '[mm] 
        dia_astap = dia_astap / 1000                                    '[m]
        lengte_astap = 1.0                                              'lengte in meters average 1m
        kg_astap = 7850 * lengte_astap * PI * (dia_astap / 2) ^ 2       'het standaardmateriaal is staal, dit is het totale inkoopmateriaal, wat daarna nog wordt gefreesd/gedraaid
        oppb_astap = PI * dia_astap * lengte_astap
        rho_kunststof = 970                                             '[kg/m3] dichtheid HDPE
        kg_lining = rho_kunststof * (PI * _diam_flight + 0.5 * (0.045 + _diam_flight / 2)) * dikte_lining * _λ6

        '---------- estimated weights---------------
        kg_inlaat = 10              '[kg] inlaat chute
        kg_uitlaat = 10             '[kg] uitlaat chute
        kg_voet = 5                 '[kg] conveyor supports
        kg_afschermkap = 10         '[kg] koppelingkap

        '---------- estimated area's---------------
        oppb_inlaat = 1             '[m2]
        oppb_uitlaat = 1            '[m2]
        oppb_voet = 0.5             '[m2]
        oppb_afschermkap = 1        '[m2]

        '----------------------------------------COST CALCULATION-----------------------------------------------
        '-------------------------------------------------------------------------------------------------------
        Dim subtotalCost_Steel As Double
        Dim subtotalCost_Components As Double
        Dim subtotalCost_Options As Double
        Dim subtotalCost_Misc As Double
        Dim kgprijs(9) As Double
        Dim totaal_gew As Double
        Dim marge_factor As Double

        'TABBLAD COSTING ---------------------------------------------------------------------------------------
        'STEEL SUBGROUP ----------------------------------------------------------------------------------------

        '======== Onderdelen van plaat die gesneden worden ==========
        gew_inuitvoet = kg_inlaat + kg_uitlaat + kg_voet + kg_afschermkap
        total_kg_plaat = 2 * kg_kopstaartplaat + kg_trog + kg_deksel + gew_inuitvoet
        tot_opperv_paint = oppb_voet + oppb_uitlaat + oppb_inlaat + oppb_trog        'Buiten Oppervlak paint onderdelen 
        totaal_gew = total_kg_plaat + kg_astap + kg_pipe + kg_schroefblad

        '+++++++++++++++++++++++ Financieel ++++++++++++++++++++++++++++++
        Double.TryParse(TextBox93.Text, kgprijs(0))     'kopstaartplaat kg prijs
        Double.TryParse(TextBox96.Text, kgprijs(1))     'trog kg prijs
        Double.TryParse(TextBox97.Text, kgprijs(2))     'deksel kg prijs
        Double.TryParse(TextBox113.Text, kgprijs(3))    'snijkosten kg prijs
        Double.TryParse(TextBox95.Text, kgprijs(4))     'schroefblad kg prijs
        Double.TryParse(TextBox94.Text, kgprijs(5))     'pijp kg prijs
        Double.TryParse(TextBox92.Text, kgprijs(6))     'astap kg prijs
        Double.TryParse(TextBox85.Text, kgprijs(7))     'Afschermkap schatting
        Double.TryParse(TextBox84.Text, kgprijs(8))     'lining kg prijs

        cost_kopstaartplaat = 2 * kg_kopstaartplaat * kgprijs(0)
        cost_trog = kg_trog * kgprijs(1)
        cost_deksel = kg_deksel * kgprijs(2)
        cost_cutting = total_kg_plaat * kgprijs(3)
        cost_schroefblad = kg_schroefblad * kgprijs(4)      'schroefblad prijs
        cost_pipe = kg_pipe * kgprijs(5)                    'pijp prijs
        cost_astap = kg_astap * kgprijs(6)                  'astap prijs

        subtotalCost_Steel = cost_kopstaartplaat + cost_trog + cost_deksel + cost_cutting + cost_schroefblad + cost_pipe + cost_astap

        '======= COMPONENTS SUBGROUP -Motorreductor-Koppeling-Lagers =======
        cost_hang = NumericUpDown35.Value * 500          'Hangende lagers
        cost_stopbus = 500                               'Asafdichtingen (2x) te ingewikkeld om precieze prijs te bepalen (2 stuks) 
        If Not CheckBox8.Checked Then cost_stopbus = 0   '-enable

        '======= SubTotal calculation =======
        subtotalCost_Components = cost_motorreductor + cost_koppeling + cost_lagers + cost_hang + cost_stopbus + cost_pakking

        'OPTIONS SUBGROUP ---------------------------------------------------------------------------------------
        cost_inlaat *= NumericUpDown20.Value            'Inlaat Schatting
        cost_uitlaat *= NumericUpDown21.Value           'Uitlaat Schatting
        cost_voet *= NumericUpDown23.Value              'Voet Schatting

        cost_afschermkap = kg_afschermkap * kgprijs(7)     'Afschermkap schatting

        'SubTotal calculation
        subtotalCost_Options = cost_inlaat + cost_uitlaat + cost_voet + cost_afschermkap

        'MISC SUBGROUP ---------------------------------------------------------------------------------------
        cost_lining = kg_lining * kgprijs(8)                 'lining 
        If Not CheckBox4.Checked Then cost_lining = 0   'enable
        cost_painting = cost_paint * tot_opperv_paint   'verf m2*prijs
        cost_transport = 400                             'intern transport
        If Not CheckBox7.Checked Then cost_transport = 0 'enable
        subtotalCost_Misc = cost_lining + cost_painting + cost_transport

        total_cost = subtotalCost_Steel + subtotalCost_Components + subtotalCost_Options + subtotalCost_Misc

        'FILL TEXTBOXES STEEL SUBGROUP ----------------------------------------------------------------------------------------
        TextBox42.Text = (2 * kg_kopstaartplaat).ToString("F1")   'kopstaartplaat kg (twee kopplaten)
        TextBox56.Text = cost_kopstaartplaat.ToString("F0")         'kopstaartplaat cost

        TextBox136.Text = NumericUpDown8.Value.ToString     'schroefblad dikte mm
        TextBox47.Text = kg_trog.ToString("F1")             'trog kg
        TextBox48.Text = kg_deksel.ToString("F1")           'deksel kg
        TextBox46.Text = kg_schroefblad.ToString("F1")      'schroefblad kg
        TextBox45.Text = kg_pipe.ToString("F1")             'pijp kg
        TextBox54.Text = kg_astap.ToString("F1")            'astap kg
        TextBox134.Text = totaal_gew.ToString("F0")         'totaal gewicht
        TextBox135.Text = gew_inuitvoet.ToString("F0")      'inlaat, uitlaat+voet

        TextBox61.Text = cost_trog.ToString("F0")           'trog cost
        TextBox62.Text = cost_deksel.ToString("F0")         'deksel cost
        TextBox43.Text = cost_cutting.ToString("F0")        'snijkosten
        TextBox60.Text = cost_schroefblad.ToString("F0")    'schroefblad cost
        TextBox59.Text = cost_pipe.ToString("F0")           'pijp cost
        TextBox78.Text = cost_astap.ToString("F0")          'astap kg


        'FILL TEXTBOXES COMPONENTS SUBGROUP ----------------------------------------------------------------------------------------
        TextBox57.Text = cost_motorreductor.ToString("F2")       'Motorreductor
        TextBox58.Text = cost_koppeling.ToString("F2")           'Koppeling
        TextBox63.Text = cost_lagers.ToString("F2")              'Lagers
        TextBox102.Text = cost_hang.ToString("F1")               'Hangende lagers
        TextBox64.Text = cost_stopbus.ToString("F1")             'Asafdichting
        TextBox104.Text = cost_pakking.ToString("F2")            'Seals

        'FILL TEXTBOXES OPTIONS SUBGROUP ---------------------------------------------------------------------------------------
        TextBox79.Text = cost_inlaat.ToString("F2")             'Inlaat cost
        TextBox80.Text = cost_uitlaat.ToString("F2")            'Uitlaat cost
        TextBox81.Text = cost_voet.ToString("F2")               'Voet cost
        TextBox82.Text = cost_afschermkap.ToString("F2")        'Afschermkap cost
        TextBox76.Text = kg_afschermkap.ToString("F0")          'Afschermkap kg

        'FILL TEXTBOXES MSIC SUBGROUP ----------------------------------------------------------------------------------------
        TextBox77.Text = kg_lining.ToString("F2")                'Lining kg
        TextBox83.Text = cost_lining.ToString("F2")              'Lining cost
        TextBox108.Text = tot_opperv_paint.ToString("F1")        'Verf m2
        TextBox107.Text = cost_painting.ToString("F2")           'Verf cost
        TextBox112.Text = cost_transport.ToString("F2")          'Transport cost
        'TABBLAD SALES PRICE ---------------------------------------------------------------------------------------

        'CALCULATE ----------------------------------------------------------------------------------------
        uren_wvb = NumericUpDown48.Value
        uren_eng = NumericUpDown30.Value
        uren_pro = NumericUpDown33.Value
        uren_fab = NumericUpDown34.Value
        tot_uren = uren_wvb + uren_eng + uren_pro + uren_fab       'Totaal aantal uren

        Dim uren_ratio(4) As Double
        uren_ratio(0) = uren_wvb / tot_uren
        uren_ratio(1) = uren_eng / tot_uren
        uren_ratio(2) = uren_pro / tot_uren
        uren_ratio(3) = uren_fab / tot_uren
        TextBox144.Text = uren_ratio(0).ToString("F2")
        TextBox145.Text = uren_ratio(1).ToString("F2")
        TextBox146.Text = uren_ratio(2).ToString("F2")
        TextBox147.Text = uren_ratio(3).ToString("F2")

        '---------- labour -------------- 
        wvb_prijs_uur = 70              'labour rate
        eng_prijs_uur = 67              'labour rate
        project_prijs_uur = 93          'labour rate
        fabriek_prijs_uur = 55          'labour rate

        prijs_wvb = uren_wvb * wvb_prijs_uur                                'Wvb cost
        prijs_eng = uren_eng * eng_prijs_uur                                'Engineering cost
        prijs_pro = uren_pro * project_prijs_uur                        'Project management cost
        prijs_fab = uren_fab * fabriek_prijs_uur                       'Fabriek cost

        tot_prijsarbeid = prijs_wvb + prijs_eng + prijs_pro + prijs_fab     'Totale prijs arbeid

        certificate_cost = 50 * NumericUpDown27.Value                       'Certificaat cost
        geheel_totprijs = total_cost + tot_prijsarbeid                      'Totaal prijs
        perc_mater = 100 * total_cost / geheel_totprijs                     'Percentage materiaal
        perc_arbeid = 100 * tot_prijsarbeid / geheel_totprijs               'Percentage arbeid
        dekking = geheel_totprijs * (1 / 0.96 - 1)                          'Risco Dekking 4%


        '------- normal customer OR intercompany -------------
        marge_factor = CDbl(IIf(RadioButton11.Checked, 0.82, 0.85))         'Marge factor
        marge_cost = (geheel_totprijs + dekking) * (1 / marge_factor - 1)   'Marge
        packing = NumericUpDown49.Value                                     'packing
        shipping = NumericUpDown50.Value                                    'Shipping
        verkoopprijs = geheel_totprijs + dekking + marge_cost               'Verkoopprijs
        verkoopprijs += packing + shipping                                  'Verkoopprijs

        'FILL TEXTBOXES ----------------------------------------------------------------------------------------
        TextBox88.Text = certificate_cost.ToString("F2")                    'Certificaat cost
        TextBox109.Text = total_kg_plaat.ToString("F0")                     'Totaal gewicht plaat
        TextBox68.Text = (cost_kopstaartplaat + cost_trog + cost_deksel + cost_cutting).ToString("F0")

        TextBox143.Text = wvb_prijs_uur.ToString("F0")              'Wvb uren
        TextBox140.Text = prijs_wvb.ToString("F0")                  'Wvb cost
        TextBox105.Text = eng_prijs_uur.ToString("F0")              'Engineering uren
        TextBox55.Text = prijs_eng.ToString("F0")                   'Engineering cost
        TextBox69.Text = project_prijs_uur.ToString("F0")           'Project management uren
        TextBox70.Text = prijs_pro.ToString("F0")                   'Project management cost
        TextBox71.Text = fabriek_prijs_uur.ToString("F0")           'Fabriek uren
        TextBox72.Text = prijs_fab.ToString("F0")                   'Fabriek cost
        TextBox106.Text = tot_uren.ToString("F0")                   'Totaal aantal uren
        TextBox111.Text = total_cost.ToString("F0")                 'Totale prijs materiaal
        TextBox103.Text = total_cost.ToString("F0")                 'Totale prijs materiaal
        TextBox100.Text = perc_mater.ToString("F0")                 'Totale percentage materiaal
        TextBox98.Text = tot_prijsarbeid.ToString("F0")             'Totale prijs arbeid
        TextBox101.Text = perc_arbeid.ToString("F0")                'Totale percentage arbeid
        TextBox73.Text = geheel_totprijs.ToString("F0")             'Geheel totaalprijs
        TextBox74.Text = dekking.ToString("F0")                     'Dekking
        TextBox99.Text = marge_cost.ToString("F0")                  'Marge
        TextBox75.Text = verkoopprijs.ToString("F0")                'Verkoopprijs
    End Sub
    Private Sub Draw_chart1()
        Dim hh As Integer

        Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()
        Chart1.Titles.Clear()

        For hh = 0 To 2
            Chart1.Series.Add("s" & hh.ToString)
            Chart1.Series(hh).ChartType = SeriesChartType.FastLine
            Chart1.Series(hh).IsVisibleInLegend = False
            Chart1.Series(hh).BorderWidth = 3
        Next
        Chart1.Series(0).Color = Color.Black
        Chart1.Series(1).Color = Color.Red

        Chart1.ChartAreas.Add("ChartArea0")
        Chart1.Series(0).ChartArea = "ChartArea0"
        Chart1.Titles.Add("Simply supported Screw conveyor" & vbCrLf & "Shear force and Bending Moment")
        Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

        '--------------- Legends and titles ---------------
        Chart1.ChartAreas("ChartArea0").AxisY.Title = "Shear force [N] and Bending Moment [N.m]"
        Chart1.ChartAreas("ChartArea0").AxisX.Title = "Shaft length [m]"
        Chart1.ChartAreas("ChartArea0").AxisY.RoundAxisValues()
        Chart1.ChartAreas("ChartArea0").AxisX.RoundAxisValues()
        Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
        Chart1.ChartAreas("ChartArea0").AxisX.Maximum = _d(_steps)

        For hh = 0 To _steps
            Chart1.Series(0).Points.AddXY(_d(hh), _s(hh)) 'Shear force line
            Chart1.Series(1).Points.AddXY(_d(hh), -_m(hh)) 'Moment line
        Next
    End Sub
    Private Sub Draw_chart2()
        Dim hh As Integer

        Chart2.Series.Clear()
        Chart2.ChartAreas.Clear()
        Chart2.Titles.Clear()

        For hh = 0 To 1
            Chart2.Series.Add("s" & hh.ToString)
            Chart2.Series(hh).ChartType = SeriesChartType.FastLine
            Chart2.Series(hh).IsVisibleInLegend = False
            Chart2.Series(hh).Color = Color.Black
            Chart2.Series(hh).BorderWidth = 2
        Next

        Chart2.ChartAreas.Add("ChartArea0")
        Chart2.Series(0).ChartArea = "ChartArea0"
        Chart2.Titles.Add("Deflection angle")
        Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

        '--------------- Legends and titles ---------------
        Chart2.ChartAreas("ChartArea0").AxisY.Title = "Deflection angle [rad]"
        Chart2.ChartAreas("ChartArea0").AxisX.Title = "Shaft length [m]"
        Chart2.ChartAreas("ChartArea0").AxisY.RoundAxisValues()
        Chart2.ChartAreas("ChartArea0").AxisX.RoundAxisValues()
        Chart2.ChartAreas("ChartArea0").AxisX.Minimum = 0
        Chart2.ChartAreas("ChartArea0").AxisX.Maximum = _d(_steps)

        For hh = 0 To _steps
            Chart2.Series(0).Points.AddXY(_d(hh), _α(hh))   'Angle
        Next
    End Sub
    Private Sub Draw_chart3()
        Dim hh As Integer

        Chart3.Series.Clear()
        Chart3.ChartAreas.Clear()
        Chart3.Titles.Clear()

        For hh = 0 To 1
            Chart3.Series.Add("s" & hh.ToString)
            Chart3.Series(hh).ChartType = SeriesChartType.FastLine
            Chart3.Series(hh).IsVisibleInLegend = False
            Chart3.Series(hh).Color = Color.Black
            Chart3.Series(hh).BorderWidth = 2
        Next

        Chart3.ChartAreas.Add("ChartArea0")
        Chart3.Series(0).ChartArea = "ChartArea0"
        Chart3.Titles.Add("Deflection in [mm]")
        Chart3.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

        '--------------- Legends and titles ---------------
        Chart3.ChartAreas("ChartArea0").AxisY.Title = "Deflection [mm]"
        Chart3.ChartAreas("ChartArea0").AxisX.Title = "Shaft length [m]"
        Chart3.ChartAreas("ChartArea0").AxisY.RoundAxisValues()
        Chart3.ChartAreas("ChartArea0").AxisX.RoundAxisValues()
        Chart3.ChartAreas("ChartArea0").AxisX.Minimum = 0
        Chart3.ChartAreas("ChartArea0").AxisX.Maximum = _d(_steps)

        For hh = 0 To _steps
            Chart3.Series(0).Points.AddXY(_d(hh), _αv(hh))  'Deflection
        Next
    End Sub
    Private Sub Screen_contrast()
        '====This fuction is to increase the readability=====
        '==== of the red text ===============================
        Dim all_txt, all_num, all_lab As New List(Of Control)

        '-------- find all Text box controls -----------------
        FindControlRecursive(all_txt, Me, GetType(TextBox))   'Find the control
        For i = 0 To all_txt.Count - 1
            Dim grbx As TextBox = CType(all_txt(i), TextBox)
            grbx.ReadOnly = False
            grbx.Enabled = True
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next

        '-------- find all numeric controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            grbx.ReadOnly = False
            grbx.Enabled = True
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next

        '-------- find all label controls -----------------
        FindControlRecursive(all_lab, Me, GetType(Label))   'Find the control
        For i = 0 To all_lab.Count - 1
            Dim grbx As Label = CType(all_lab(i), Label)
            grbx.Enabled = True
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next
    End Sub
End Class
