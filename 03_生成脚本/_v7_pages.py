# ================================================================ 01 封面
s = newslide()
box(s, 0, 0, W, H, C['WHITE'])
box(s, 0, 0, W, 6, C['NAVY'])
box(s, 0, 0, 320, 6, C['LIME'])
box(s, 0, 6, 6, H - 6, C['NAVY'])
tx(s, X0 + 20, 96, 460, 20, 'HYTERA  \u00d7  GLOBE TELECOM', 11, True, 'GREEN')
tx(s, X0 + 20, 126, 470, 90, '450 MHz Mission-Critical\nNetwork for the Philippines', 31, True, 'NAVY', line=1.12)
box(s, X0 + 20, 232, 66, 4, C['LIME'])
tx(s, X0 + 20, 252, 460, 40,
   'A coverage and resilience layer for utilities, public safety, transport and remote communities.',
   12.5, False, 'STEEL', line=1.35)
for i, (t, ic) in enumerate([('Terminals', '\u7ebf\u7a3f_\u624b\u53f0_2.png'),
                             ('Base stations', '\u901a\u4fe1\u94c1\u5854_3.png'),
                             ('Dispatch platform', '\u7ebf\u7a3f_\u7ec8\u7aef_1.png')]):
    x = X0 + 20 + i * 152
    box(s, x, 304, 140, 56, C['CARD'])
    fit(s, illu(ic), x + 10, 314, 26, 36)
    tx(s, x + 44, 304, 92, 56, t, 9.5, True, 'NAVY', L, MID, line=1.2)
box(s, X0 + 20, 376, 460, 34, C['PAPER'])
tx(s, X0 + 36, 376, 440, 34, 'Philippines  \u00b7  October 2026', 11.5, True, 'NAVY', L, MID)

box(s, 540, 40, 396, 370, C['PAPER'])
fit(s, illu('\u83f2\u5f8b\u5bbe\u89e3\u51b3\u65b9\u6848\u5305\u88c5_slide11_1.png', 1400), 556, 56, 130, 338)
fit(s, illu('Hytera 450M Product Introduction PPT_202609_slide9_5.png'), 700, 84, 220, 176)
tx(s, 700, 274, 220, 60,
   'One network, six industries:\npower, water, police, emergency,\nrail and port.', 10.5, True, 'NAVY', line=1.35)
box(s, 540, 40, 396, 370, None, C['LINE'], 1.0)

box(s, 0, 424, W, 116, C['CARD'])
box(s, 0, 420, W, 4, C['LIME'])
for i, n in enumerate(['\u901a\u4fe1\u94c1\u5854_3.png', '\u57ce\u5e02\u697c\u7fa4.png',
                       '\u6c34\u5e93\u5927\u575d.png', '\u4ea4\u901a\u67a2\u7ebd.png',
                       '\u8239\u8236.png', '\u96c6\u88c5\u7bb1\u6307\u6325\u6240.png',
                       '\u5de5\u5382.png']):
    fit(s, illu(n), 62 + i * 122, 434, 104, 96)
tx(s, X0, 528, 700, 12, 'Hytera Communications Corporation Limited', 8, False, 'SILVER')

# ================================================================ 02 国情
s = newslide()
page(s, 2, 'The gap', 'Islands, mountains and typhoons',
     'Public 4G and 5G reach the cities. What is missing is coverage that is still there after the storm.')
box(s, X0, 122, 400, 296, C['PAPER'])
fit(s, illu('\u83f2\u5f8b\u5bbe\u89e3\u51b3\u65b9\u6848\u5305\u88c5_slide11_1.png', 1400), X0 + 20, 138, 170, 264)
fit(s, illu('\u5c71\u8109\u7ebf\u7a3f.png'), X0 + 196, 250, 190, 56)
fit(s, illu('\u6811\u6728\u526a\u5f71.png'), X0 + 196, 310, 80, 90)
fit(s, illu('\u4eba\u7269_\u5de5\u4eba.png'), X0 + 290, 296, 60, 104)
box(s, X0 + 196, 138, 190, 96, C['CARD'])
tx(s, X0 + 212, 152, 158, 70, '7,641 islands.\nOne grid has to reach\nacross all of them.',
   10.5, True, 'NAVY', line=1.35)
STATS = [('7,641', 'islands, much of the terrain mountainous', 'NAVY'),
         ('#1', 'WorldRiskIndex 2025 - highest disaster risk', 'GREEN'),
         ('~20', 'tropical cyclones enter PAR every year', 'NAVY'),
         ('1.4M', 'households lost power in Typhoon Tino', 'GREEN')]
for i, (v, lab, ac) in enumerate(STATS):
    r, c = divmod(i, 2)
    stat(s, 480 + c * 220, 122 + r * 152, 208, 140, v, lab, ac, 24)
concl(s, [('450 MHz is the layer that stays up.  ', True, 'WHITE'),
          ('Coverage and resilience, not capacity - alongside Globe\u2019s 4G and 5G, not instead of them.',
           False, 'MIST')])
src(s, 'WorldRiskReport 2025; PAGASA; Philippine Statistics Authority')

# ================================================================ 03 目录
s = newslide()
page(s, 3, 'Agenda', 'What we will cover',
     'Five sections. The first two are the argument, the last three are the delivery.')
AG = [('01', 'Why 450 MHz & Hytera', 'Global trend and our position',
       '\u901a\u4fe1\u94c1\u5854_3.png'),
      ('02', 'Industry solutions', 'Power \u00b7 Water \u00b7 Police \u00b7 Emergency \u00b7 Rail \u00b7 Port',
       '\u4ea4\u901a\u67a2\u7ebd.png'),
      ('03', 'End-to-end architecture', 'Terminal \u00b7 Network \u00b7 Application',
       'GSM\u7ec4\u7f51\u793a\u610f.png'),
      ('04', 'Reference cases', 'Global 450 MHz networks and MCX at scale',
       '\u673a\u573a.png' if False else '\u529e\u516c\u56ed\u533a.png'),
      ('05', 'Business model', 'Who does what, revenue streams, next steps',
       '\u529e\u516c\u697c.png')]
for i, (num, title, sub, ic) in enumerate(AG):
    x = X0 + i * 174
    box(s, x, 122, 162, 296, C['CARD'])
    box(s, x, 122, 162, 118, C['PAPER'])
    fit(s, illu(ic), x + 21, 130, 120, 102)
    tx(s, x + 14, 250, 134, 30, num, 21, True, 'LIME')
    tx(s, x + 14, 284, 134, 40, title, 11.5, True, 'NAVY', line=1.2)
    tx(s, x + 14, 328, 134, 70, sub, 9, False, 'STEEL', line=1.3)
concl(s, [('One network, six industries.  ', True, 'WHITE'),
          ('The same 450 MHz layer serves all of them, on the same sites.', False, 'MIST')])
src(s, 'Hytera proposal structure')

# ================================================================ 04 450MHz 全球
s = newslide()
page(s, 4, 'Global trend', 'From a niche band to critical infrastructure',
     'Regulators, utilities and public safety agencies have settled on what this band is for.')
TREND = [('86%', 'of the world\u2019s population lives in countries where the 400 MHz bands are allocated or under consultation', 'NAVY'),
         ('39', 'countries have already deployed networks in these bands, out of 77 with allocations at 380 / 410 / 450 MHz', 'GREEN'),
         ('1.8M / 14M', 'broadband and sensor connections forecast by 2030, from 420k and 1.2M in 2026', 'NAVY')]
for i, (v, lab, ac) in enumerate(TREND):
    stat(s, X0, 122 + i * 86, 470, 78, v, lab, ac, 21)
POS = [('\u7ebf\u7a3f_\u624b\u53f0_2.png', 'Named at both ends of the chain',
        'The Alliance lists Hytera in its device supplier list and in its macro eNB supplier list.'),
       ('\u901a\u4fe1\u94c1\u5854_3.png', 'Through operator acceptance',
        'PNC660 450M is whitelisted by 450connect to \u2212115 dBm RSRP - the operator\u2019s own document.'),
       ('\u7ebf\u7a3f_\u7ec8\u7aef_1.png', 'A 5G-ready 450 MHz portfolio',
        'Handheld, vehicle radio, base station and core, with a 5G model in development.')]
klab(s, 546, 122, 'Our position')
for i, (ic, t, d) in enumerate(POS):
    y = 142 + i * 84
    box(s, 546, y, 366, 76, C['CARD'])
    box(s, 546, y, 3.5, 76, C['NAVY'] if i == 0 else C['MIST'])
    fit(s, illu(ic), 560, y + 14, 22, 48)
    tx(s, 592, y + 12, 308, 18, t, 11, True, 'NAVY')
    tx(s, 592, y + 32, 308, 40, d, 9, False, 'STEEL', line=1.3)
concl(s, [('The band is no longer experimental.  ', True, 'WHITE'),
          ('It is assigned, standardised in 3GPP, and carrying utility and public-safety traffic today.',
           False, 'MIST')])
src(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p4, p11, p12; Annual Equipment Report (May 2025) p5')

# ================================================================ 05 物理层面
s = newslide()
page(s, 5, 'Why 450 MHz', 'One site replaces several mid-band sites',
     'Low-band physics is the whole argument. The numbers below reuse one grid at three frequencies.')
DATA = [('450 MHz', '1 site  \u00b7  100% covered', [(0, 0)], 62),
        ('740 MHz', '2.5-3\u00d7 sites  \u00b7  91.3% from the same grid',
         [(-36, -16), (36, -16), (0, 28)], 34),
        ('2350 MHz', '12-15\u00d7 sites  \u00b7  58.7% from the same grid',
         [(i * 38 - 57, j * 30 - 45) for i in range(4) for j in range(4)
          if (i, j) not in ((0, 0), (3, 0), (0, 3))], 17)]
for i, (name, sub, sites, r) in enumerate(DATA):
    x = X0 + i * 166
    box(s, x, 122, 156, 156, C['PAPER'])
    for k in range(1, 4):
        ln(s, x + k * 39, 127, x + k * 39, 273, C['MIST'], 0.5)
        ln(s, x + 5, 122 + k * 39, x + 151, 122 + k * 39, C['MIST'], 0.5)
    cx, cy = x + 78, 200
    for dx, dy in sites:
        box(s, cx + dx - r, cy + dy - r * 0.86, 2 * r, 2 * r * 0.86,
            'D8E4F6', C['NAVY'], 0.75, shape=MSO_SHAPE.OVAL)
    fit(s, illu('\u901a\u4fe1\u94c1\u5854_2.png'), x + 62, 150, 30, 62)
    tx(s, x, 288, 156, 16, name, 11, True, 'NAVY', CEN)
    tx(s, x, 306, 156, 30, sub, 9, True, 'GREEN', CEN, line=1.25)
tx(s, X0, 344, 500, 32,
   'Same area, same grid. Only the low band covers it with one site - which is what makes island and\n'
   'mountain coverage affordable, and what makes extensive backup power affordable.',
   10, False, 'STEEL', line=1.35)
klab(s, 560, 122, 'What that buys')
bullets(s, 560, 148, 352, [
    'Bends around ridges and into valleys instead of needing line of sight',
    'Reaches meter boxes, basements and tunnels that mid-band cannot',
    'Fewer sites is what makes extensive backup power affordable',
], 11.5, 50)
concl(s, [('What it is not:  ', True, 'WHITE'),
          ('with 3-5 MHz it carries voice, low-rate video and telemetry. It stays outside the capacity game.',
           False, 'MIST')])
src(s, '450 MHz Alliance, Brief (Jun 2022) p15')

# ================================================================ 06 MCS
s = newslide()
page(s, 6, 'The service layer', 'What MCS means - and how it sits on the band',
     'Every case in this deck runs MCS. This page explains what it is, because the term carries the delivery.')
box(s, X0, 122, 430, 200, C['CARD'])
tx(s, X0 + 18, 136, 394, 18, '3GPP mission-critical services (MCS / MCX)', 12, True, 'NAVY')
MCS = [('\u7ebf\u7a3f_\u624b\u53f0_2.png', 'MCPTT', 'push-to-talk voice, with priority and pre-emption'),
       ('\u7ebf\u7a3f_\u624b\u53f0_3.png', 'MCVideo', 'live and recorded video from the scene'),
       ('\u7ebf\u7a3f_\u7ec8\u7aef_1.png', 'MCData', 'messaging, file transfer and location data')]
for i, (ic, k, d) in enumerate(MCS):
    y = 164 + i * 48
    box(s, X0 + 18, y, 394, 40, C['WHITE'])
    box(s, X0 + 18, y, 3, 40, C['LIME'])
    fit(s, illu(ic), X0 + 30, y + 6, 16, 28)
    tx(s, X0 + 54, y + 4, 110, 18, k, 11.5, True, 'NAVY')
    tx(s, X0 + 54, y + 21, 344, 16, d, 9, False, 'STEEL')
tx(s, X0, 330, 430, 44,
   'Defined by 3GPP, not by a single vendor - so the network stays multi-vendor and the\n'
   'terminals are not locked to one supplier.', 10, False, 'STEEL', line=1.35)
STACK = [('Applications', 'dispatch, recording, device management', 'NAVY'),
         ('MCS service layer', 'MCPTT / MCVideo / MCData', 'GREEN'),
         ('LTE or 5G radio', 'Band 31 / 72 at 450 MHz', 'NAVY_L')]
klab(s, 510, 122, 'How the layers stack')
for i, (t, d, col) in enumerate(STACK):
    y = 146 + i * 62
    box(s, 510, y, 402, 52, C[col])
    tx(s, 528, y + 8, 366, 18, t, 12, True, 'WHITE')
    tx(s, 528, y + 28, 366, 16, d, 9, False, 'MIST')
    if i < 2:
        box(s, 700, y + 54, 22, 16, C['NAVY_L'], shape=MSO_SHAPE.DOWN_ARROW)
box(s, 510, 336, 402, 38, C['GREEN_L'])
tx(s, 528, 336, 366, 38,
   'In this deck: the band carries the traffic, MCS provides the service.',
   10, True, 'GREEN', L, MID)
concl(s, [('450 MHz gives the coverage; MCS gives the service.  ', True, 'WHITE'),
          ('The two are specified together, and the cases that follow run both.', False, 'MIST')])
src(s, '3GPP mission-critical services; Hytera HyTalk MC platform documentation')

# ================================================================ 07 六行业总览
s = newslide()
page(s, 7, 'Industry solutions', 'One network, six industries',
     'The same 450 MHz layer serves all six, on the same sites. Each of the next pages takes one.')
IND = [('Power', 'Grid dispatch, feeder automation and smart metering', '\u7535\u529b\u62a2\u4fee.png'),
       ('Water', 'Metering, pressure sensing and reservoir telemetry', '\u6c34\u5e93\u5927\u575d.png'),
       ('Police', 'Cross-agency group calls and live video from the scene', '\u8b66\u5bdf\u5c40.png'),
       ('Emergency', 'Rapid deployment and coverage after the storm', '\u6d88\u9632\u7ad9\u706b\u707e.png'),
       ('Rail', 'Corridor coverage and train-to-ground voice and data', '\u4ea4\u901a\u67a2\u7ebd.png'),
       ('Port', 'Quay dispatch, ship-to-shore and coast guard patrols', '\u8239\u8236.png')]
for i, (t, d, ic) in enumerate(IND):
    r, c = divmod(i, 3)
    x = X0 + c * 292
    y = 122 + r * 152
    box(s, x, y, 280, 140, C['CARD'])
    box(s, x, y, 3.5, 140, C['NAVY'] if r == 0 else C['LIME'])
    box(s, x + 14, y + 12, 116, 116, C['PAPER'])
    fit(s, illu(ic), x + 20, y + 18, 104, 104)
    tx(s, x + 144, y + 24, 124, 22, t, 14.5, True, 'NAVY')
    tx(s, x + 144, y + 54, 124, 72, d, 9.5, False, 'STEEL', line=1.35)
concl(s, [('Utilities anchor it commercially.  ', True, 'WHITE'),
          ('Public safety, emergency, rail and port add national value on the same sites.', False, 'MIST')])
src(s, '450 MHz Alliance (Dec 2025); Hytera industry solutions')

# ================================================================ 08-13 六个行业页
IND_PAGES = [
    dict(n=8, kicker='Industry 1 of 6  \u00b7  Power', title='Power: the grid, the crews and the meters',
         point='Electric cooperatives have spread-out assets, a 10-15 year horizon and a willingness to pay for reliability.',
         hero='\u7535\u529b\u62a2\u4fee.png', accent='\u901a\u4fe1\u94c1\u5854_1.png',
         cap='Grid repair crews and line assets. Illustrative imagery.',
         pain='Mountain and island lines have no reliable signal, and feeder automation needs a private data path that does not depend on the public network.',
         sol=['Wide-area push-to-talk, video and lone-worker alarms along the lines',
              'A private data channel for SCADA, feeder automation and substation control',
              'Smart meters on the same sites, with deep indoor coverage'],
         ph='In the Philippines: 121 electric cooperatives under NEA.',
         srcn='450 MHz Alliance (Dec 2025); NEA; Hytera power industry solutions'),
    dict(n=9, kicker='Industry 2 of 6  \u00b7  Water', title='Water: locate the losses instead of estimating them',
         point='Non-revenue water is usually estimated, not located. Pump, well and reservoir sites sit exactly where mid-band cannot reach.',
         hero='\u6c34\u5e93\u5927\u575d.png', accent='\u5c71\u8c37\u6cb3\u6d41.png',
         cap='Reservoir and pumping assets. Illustrative imagery.',
         pain='Most water districts know their losses as a percentage, not as an address - because the meters and pressure sensors cannot report.',
         sol=['Meters and pressure sensors that locate losses instead of estimating them',
              'Pump, well and reservoir telemetry from remote and underground sites',
              'Backup-powered sites keep alarms flowing when the public network is down'],
         ph='In the Philippines: 149 water districts under LWUA.',
         srcn='450 MHz Alliance (Dec 2025); LWUA; Hytera water industry solutions'),
    dict(n=10, kicker='Industry 3 of 6  \u00b7  Police', title='Police: one group call across the agencies',
         point='The requirement is not another private network. It is interoperability between the agencies that already have one.',
         hero='\u8b66\u5bdf\u5c40.png', accent='\u8b66\u8f66_1.png',
         cap='A district police station. Illustrative imagery.',
         pain='Different agencies cannot talk to each other, and the public network congests exactly at the moment it is needed most.',
         sol=['One group call across police, fire and civil defence',
              'Live video from the scene, with priority and pre-emption when the network congests',
              'Interworking with the DMR and TETRA networks already in service'],
         ph='Reference: Singapore has assigned 450 MHz band 31 to public safety, in service.',
         srcn='450 MHz Alliance, Annual Global Update p8; Hytera public-safety references'),
    dict(n=11, kicker='Industry 4 of 6  \u00b7  Emergency', title='Emergency: coverage that survives the storm',
         point='After a typhoon the public network fails exactly where the response is needed. This is the layer that has to survive.',
         hero='\u6d88\u9632\u7ad9\u706b\u707e.png', accent='\u57ce\u5e02\u5185\u6d9d.png',
         cap='Fire, flooding and road clearance. Illustrative imagery.',
         pain='Rapid response needs communications within hours, not days - and it needs them where there is no grid and no backhaul.',
         sol=['Portable base stations and satellite backhaul deployed within hours',
              'A field command post that reaches the provincial EOC and the national agency',
              'Sites that keep running on backup power when the grid does not'],
         ph='Reference: rapid deployment kits and container command posts in service.',
         srcn='Hytera emergency and rapid-deployment product documentation'),
    dict(n=12, kicker='Industry 5 of 6  \u00b7  Rail', title='Rail: continuous coverage along a narrow corridor',
         point='A rail corridor is a narrow strip, often in cuttings. It is one of the two shapes mid-band handles worst.',
         hero='\u4ea4\u901a\u67a2\u7ebd.png', accent='\u96a7\u9053.png',
         cap='An interchange and a tunnel section. Illustrative imagery.',
         pain='Coverage has to follow the track, not the population - so a network built for towns leaves the corridor uncovered.',
         sol=['Continuous coverage along the corridor, including cuttings and tunnels',
              'Train-to-ground voice, data and video for operations and security',
              'Runs on the operator\u2019s network rather than a separate private system'],
         ph='Reference: South Africa has assigned the band to Transnet for rail.',
         srcn='450 MHz Alliance, Annual Global Update 2025 p6; Hytera railway references'),
    dict(n=13, kicker='Industry 6 of 6  \u00b7  Port', title='Port: from the quay out to sea',
         point='Ports are dense and metallic, and coast guard patrols run far beyond mid-band reach. One network covers both.',
         hero='\u8239\u8236.png', accent='\u96c6\u88c5\u7bb1\u6307\u6325\u6240.png',
         cap='A coastal terminal and a container command post. Illustrative imagery.',
         pain='Quay operations, ship-to-shore traffic and patrol boats all need the same group call, across water and steel.',
         sol=['Port and terminal dispatch on one group call across the quay',
              'Ship-to-shore voice and data from coastal sites',
              'Coast guard patrols covered far out to sea, on the same network as the port'],
         ph='The same coastal sites also serve island and rural communities.',
         srcn='450 MHz Alliance (Dec 2025); Hytera transport and maritime references'),
]
for cfg in IND_PAGES:
    s = newslide()
    page(s, cfg['n'], cfg['kicker'], cfg['title'], cfg['point'])
    box(s, X0, 122, 380, 230, C['PAPER'])
    fit(s, illu(cfg['hero']), X0 + 34, 140, 250, 194)
    fit(s, illu(cfg['accent']), X0 + 296, 196, 66, 84)
    box(s, X0, 122, 380, 230, None, C['LINE'], 1.0)
    tx(s, X0, 356, 380, 30, cfg['cap'], 8.5, False, 'SILVER', line=1.3)
    box(s, 456, 122, 456, 78, C['GREEN_L'])
    box(s, 456, 122, 3.5, 78, C['GREEN'])
    tx(s, 474, 132, 424, 14, 'PAIN POINT', 8.5, True, 'GREEN')
    tx(s, 474, 150, 424, 46, cfg['pain'], 9.5, False, 'STEEL', line=1.3)
    klab(s, 456, 210, 'What 450 MHz does')
    for j, it in enumerate(cfg['sol']):
        yy = 232 + j * 40
        fit(s, illu('\u7ebf\u7a3f_\u7ec8\u7aef_1.png'), 456, yy + 1, 12, 20)
        tx(s, 478, yy, 434, 36, it, 10.5, False, 'STEEL', line=1.28)
    box(s, 456, 356, 456, 30, C['CARD'])
    tx(s, 472, 356, 424, 30, cfg['ph'], 9.5, True, 'NAVY', L, MID)
    src(s, cfg['srcn'])

# ================================================================ 14 端到端架构
s = newslide()
page(s, 14, 'Architecture', 'End to end: terminal - network - application',
     'One architecture, in three layers. Each layer is described by what it contains, not by who supplies it.')
LAY = [('TERMINAL', 'What the field carries', '\u7ebf\u7a3f_\u624b\u53f0_2.png',
        [('PNC660 450M handheld', 'B31/B72, IP68, dual SIM'),
         ('MNC650 450M', 'vehicle radio'),
         ('P70 5G+450M', 'radio-style handheld, 2026 Q4'),
         ('PNC660Ex', 'ATEX variant, 2027 Q2')],
        'Hytera 450M Product Introduction PPT_202609_slide9_5.png'),
       ('NETWORK', 'Radio, core and backhaul', '\u901a\u4fe1\u94c1\u5854_2.png',
        [('450 MHz macro eNB', 'wide area coverage, few sites'),
         ('Private core', 'or the operator\u2019s own core'),
         ('Backhaul', 'fibre, microwave or satellite'),
         ('Rapid deployment', 'portable base station and satellite kit')],
        'iBS\u4e00\u4f53\u5316\u57fa\u7ad9_1.png'),
       ('APPLICATION', 'Dispatch, recording and management', '\u7ebf\u7a3f_\u7ec8\u7aef_1.png',
        [('MCS services', 'MCPTT / MCVideo / MCData'),
         ('Dispatch console', 'voice, video, GIS and presence'),
         ('Recording', 'multimedia capture and playback'),
         ('Device management', 'zero-touch deployment')],
        'Hytera PNC660 450MHz&MCX_slide8_3.png')]
for i, (name, sub, ic, items, img) in enumerate(LAY):
    x = X0 + i * 296
    box(s, x, 122, 280, 250, C['CARD'])
    box(s, x, 122, 280, 36, C['NAVY'])
    tx(s, x + 44, 122, 220, 36, name, 13, True, 'WHITE', L, MID)
    box(s, x + 14, 128, 24, 24, C['WHITE'])
    fit(s, illu(ic), x + 18, 131, 16, 18)
    tx(s, x + 16, 166, 248, 16, sub, 10, True, 'GREEN')
    for j, (k, d) in enumerate(items):
        yy = 190 + j * 34
        box(s, x + 16, yy + 5, 5, 5, C['LIME'])
        tx(s, x + 28, yy, 240, 14, k, 9.5, True, 'NAVY')
        tx(s, x + 28, yy + 14, 240, 14, d, 8.5, False, 'STEEL')
    box(s, x + 14, 322, 252, 40, C['WHITE'])
    fit(s, illu(img), x + 20, 326, 240, 32)
    if i < 2:
        box(s, x + 282, 232, 12, 26, C['NAVY_L'], shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
box(s, X0, 382, CW, 52, C['PAPER'])
tx(s, X0 + 18, 390, 300, 16, 'THE SAME ARCHITECTURE SERVES ALL SIX INDUSTRIES', 8.5, True, 'GREEN')
for i, t in enumerate(['Power', 'Water', 'Police', 'Emergency', 'Rail', 'Port']):
    x = X0 + 18 + i * 140
    box(s, x, 408, 130, 20, C['CARD'])
    tx(s, x, 408, 130, 20, t, 9, True, 'NAVY', CEN, MID)
concl(s, [('Only the terminals, the applications and the coverage area change.  ', True, 'WHITE'),
          ('The network layer is the same build.', False, 'MIST')])
src(s, '450M Roadmap; HyTalk MC V6.1; 450 MHz Alliance, Equipment Report (May 2025)')

# ================================================================ 15 全球运营商
s = newslide()
page(s, 15, 'Reference  \u00b7  450 MHz', '450 MHz operators worldwide: the model works',
     'Two things are the same in every market: a licensed operator carries the traffic, and utilities anchor the demand.')
box(s, X0, 122, 400, 296, C['NAVY'])
fit(s, illu('alliance_worldmap.png', 1400), X0 + 14, 134, 372, 258)
tx(s, X0 + 14, 396, 372, 16, '400 MHz band allocations and deployments worldwide', 9, True, 'WHITE')
OPS = [('Germany', '450connect', 'Band 72, nationwide utility network; more than 95% of Germany covered'),
       ('Austria', 'ArgoNET', 'Band 72; regional CDMA and LTE networks for utilities'),
       ('Poland', 'PGE / Polkomtel', 'Band 31 and band 87; smart grid and field services'),
       ('Malaysia', 'TNB', 'Band 31; procurement for initial deployment in four regions'),
       ('Indonesia', 'SeaTelSat', 'Rural connectivity and utility projects'),
       ('South Africa', 'Transnet', 'Band 31 allocated for rail, with wider use under discussion'),
       ('Singapore', 'Public safety', 'Band 31 allocated for public safety, in service')]
for i, (c, op, d) in enumerate(OPS):
    y = 122 + i * 42
    box(s, 468, y, 444, 38, C['CARD'] if i % 2 == 0 else C['PAPER'])
    tx(s, 480, y + 3, 96, 16, c, 9.5, True, 'NAVY')
    tx(s, 480, y + 20, 96, 14, op, 8, True, 'GREEN')
    tx(s, 584, y + 4, 320, 30, d, 8.5, False, 'STEEL', line=1.25)
concl(s, [('The pattern is the same everywhere.  ', True, 'WHITE'),
          ('A licensed operator carries the traffic; utilities and public safety anchor the demand.', False, 'MIST')])
src(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p7-p9, p14; RFI for devices (2025) p5-p7')

# ================================================================ 16-18 MCS 案例
CASES = [
    dict(n=16, kicker='Case 1 of 3  \u00b7  Railway MCX',
         title='Sri Lanka Railways: a national network on public LTE',
         point='Nine lines and 403 stations, on the operator\u2019s public LTE network rather than a separate private system.',
         img=asset('Hytera PNC660 450MHz&MCX_slide13_1.jpg'), fy=0.5,
         cap='Sri Lanka Railways - the customer\u2019s own network, photographed on site.',
         stats=[('1,950+', 'MCS terminals delivered across the island network'),
                ('9 / 403', 'lines and stations covered, in three provinces'),
                ('MCX', '3GPP mission-critical voice and data, on public LTE')],
         metrics=[('\u221210%', 'train delays'), ('\u221260%', 'management cost'),
                  ('\u221240%', 'track repair time'), ('30 min', 'landslide response')],
         srcn='Hytera showroom briefing (SRC-0134/0135) L211; Hytera PNC660 450MHz&MCX case records'),
    dict(n=17, kicker='Case 2 of 3  \u00b7  Railway MCX over 5G',
         title='Hong Kong Light Rail: the first MCS over 5G',
         point='The world\u2019s first railway MCS over 5G - replacing TETRA rather than overlaying it.',
         img=enhance('v6_dispatch.png', asset('20260917_\u4ea7\u54c1\u6587\u6863_SmartDispatchPlus\u4ea7\u54c1\u4ecb\u7ecd_\u8bfe\u4ef6_\u82f1\u6587_slide5_6.png')),
         fy=0.5,
         cap='Illustrative operations imagery; the customer\u2019s own site photography is not cleared for external use.',
         stats=[('810', 'terminals: 511 handheld, 152 bus-mounted and 147 vehicle'),
                ('5G', 'the world\u2019s first railway MCS running over 5G'),
                ('Full', 'redundancy, including the wireless links between sites')],
         metrics=[('TETRA', 'replaced, not overlaid'), ('Shared', 'operator core network and RAN'),
                  ('Train-borne', 'MCS-compliant dispatch'), ('One', 'network for operations and security')],
         srcn='Hytera showroom briefing L170; Hytera PNC660 450MHz&MCX case records'),
    dict(n=18, kicker='Case 3 of 3  \u00b7  Public safety',
         title='Papua New Guinea: a national command centre',
         point='A police headquarters command centre on Hytera ICC, used for daily operations and for APEC summit security.',
         img=asset('20260903_\u6d77\u672c\u4e8c\u90e8\u57f9\u8bad_HyTalk\u4e0eHyTalk-Pro\u65b9\u6848_\u8bfe\u4ef6_20260905_slide59_3.jpg'),
         fy=0.5,
         cap='Illustrative operations imagery; the customer\u2019s own site photography is not cleared for external use.',
         stats=[('3,000', 'police officers and 600 police vehicles served by one centre'),
                ('APEC', 'summit security: 15,000 foreign visitors and more than 200 meetings'),
                ('360\u00b0', 'city-wide video surveillance, with protection around the clock')],
         metrics=[('60 s', 'police dispatch after the alarm bell rings'),
                  ('VIP', 'automatic vehicle route monitoring'),
                  ('City-wide', 'emergency video communication'),
                  ('24/7', 'protection during the summit')],
         srcn='Hytera case records (Papua New Guinea Police HQ command centre), slides 2-4'),
]
for cfg in CASES:
    s = newslide()
    page(s, cfg['n'], cfg['kicker'], cfg['title'], cfg['point'])
    photo(s, cfg['img'], X0, 122, 470, 246, fy=cfg['fy'])
    box(s, X0, 122, 470, 246, None, C['LINE'], 1.0)
    tx(s, X0, 374, 470, 24, cfg['cap'], 8.5, False, 'SILVER', line=1.3)
    for i, (v, lab) in enumerate(cfg['stats']):
        y = 122 + i * 86
        box(s, 558, y, 354, 76, C['CARD'])
        box(s, 558, y, 3.5, 76, C['NAVY'] if i == 0 else C['MIST'])
        tx(s, 576, y + 10, 140, 28, v, 18, True, 'NAVY')
        tx(s, 576, y + 40, 320, 30, lab, 9.5, False, 'STEEL', line=1.25)
    box(s, X0, 406, CW, 34, C['NAVY'])
    for i, (v, lab) in enumerate(cfg['metrics']):
        x = X0 + 12 + i * 212
        tx(s, x, 409, 204, 18, v, 12, True, 'LIME')
        tx(s, x, 426, 204, 12, lab, 8, False, 'MIST')
    src(s, cfg['srcn'])

# ================================================================ 19 商业模式
s = newslide()
page(s, 19, 'Business model', 'Who does what',
     'Three parties, one chain. The question is not whether the technology works, but who commits first.')
FLOW = [('Hytera', 'supplies the radio, terminals, platform and integration',
         '\u5de5\u5382.png', 'NAVY'),
        ('Globe Telecom', 'holds the spectrum, builds and operates the network, owns the customer relationship',
         '\u901a\u4fe1\u94c1\u5854_3.png', 'GREEN'),
        ('Anchor customers', 'utilities, public safety, transport and ports contract capacity before build-out',
         '\u529e\u516c\u56ed\u533a.png', 'NAVY_L')]
for i, (t, d, ic, col) in enumerate(FLOW):
    x = X0 + i * 300
    box(s, x, 122, 264, 116, C[col])
    box(s, x + 16, 134, 44, 44, C['WHITE'])
    fit(s, illu(ic), x + 20, 138, 36, 36)
    tx(s, x + 72, 138, 178, 20, t, 14, True, 'WHITE')
    tx(s, x + 16, 186, 232, 44, d, 9.5, False, 'MIST', line=1.3)
    if i < 2:
        box(s, x + 268, 166, 26, 34, C['LIME'], shape=MSO_SHAPE.RIGHT_ARROW)
klab(s, X0, 256, 'Three ways to cut the cost')
COST = [('\u901a\u4fe1\u94c1\u5854_2.png', 'Reuse what Globe already has',
         'Towers, fibre, equipment rooms and power are already in place along the routes that matter.'),
        ('\u7ebf\u7a3f_\u624b\u53f0_2.png', 'Far fewer sites for the same coverage',
         'The site count for the same area is a fraction of mid-band - which is what makes backup power affordable.'),
        ('\u529e\u516c\u697c.png', 'Contract the demand before the build',
         'Anchor customers commit capacity first, so the network is built against revenue rather than in front of it.')]
for i, (ic, t, d) in enumerate(COST):
    x = X0 + i * 292
    box(s, x, 280, 280, 100, C['CARD'])
    box(s, x, 280, 3.5, 100, C['LIME'])
    fit(s, illu(ic), x + 16, 294, 28, 40)
    tx(s, x + 54, 292, 210, 22, t, 11, True, 'NAVY', line=1.2)
    tx(s, x + 16, 336, 248, 40, d, 9, False, 'STEEL', line=1.32)
box(s, X0, 392, CW, 44, C['PAPER'])
box(s, X0, 392, 3.5, 44, C['GREEN'])
rich(s, X0 + 18, 392, CW - 36, 44,
     [('The Alliance\u2019s own assessment:  ', True, 'NAVY'),
      ('the main risk in this band is commercial uptake, not deployment - which is why the sequencing matters more than the technology.',
       False, 'STEEL')], 11, anchor=MID)
src(s, '450 MHz Alliance, Annual Global Update (Dec 2025); Hytera proposal')

# ================================================================ 20 收入
s = newslide()
page(s, 20, 'Business model', 'Six revenue streams on one network',
     'A single 450 MHz build supports six commercial lines. None of them depends on consumer take-up.')
REV = [('\u7ebf\u7a3f_\u7ec8\u7aef_1.png', 'IoT connectivity', 'Monthly fee per meter, sensor or asset, on the same sites as the voice service.'),
       ('\u7ebf\u7a3f_\u624b\u53f0_2.png', 'MCX as a service', 'Hosted push-to-talk, video and data, charged per user per month.'),
       ('\u901a\u4fe1\u94c1\u5854_2.png', 'Capacity packages', 'Reserved network share for an anchor customer, on a fixed annual fee.'),
       ('\u624b\u673a.png' if False else '\u7ebf\u7a3f_\u6267\u6cd5\u8bb0\u5f55\u4eea.png', 'Device as a service', 'Terminals bundled with the service, billed monthly rather than bought outright.'),
       ('\u529e\u516c\u697c.png', 'Integration and O&M', 'SCADA, telemetry and smart-city projects, delivered and maintained.'),
       ('\u79fb\u52a8\u65b9\u8231.png', 'Emergency readiness', 'Rapid deployment kit on standby, retained by the province or the agency.')]
for i, (ic, t, d) in enumerate(REV):
    r, c = divmod(i, 3)
    x = X0 + c * 292
    y = 122 + r * 156
    box(s, x, y, 280, 144, C['CARD'])
    box(s, x, y, 280, 3.5, C['NAVY'] if r == 0 else C['LIME'])
    box(s, x + 16, y + 18, 46, 46, C['WHITE'])
    fit(s, illu(ic), x + 22, y + 24, 34, 34)
    tx(s, x + 74, y + 24, 190, 36, t, 12.5, True, 'NAVY', line=1.2)
    tx(s, x + 16, y + 78, 248, 56, d, 9.5, False, 'STEEL', line=1.35)
concl(s, [('One build, six lines.  ', True, 'WHITE'),
          ('The utility contract funds the coverage that the other five then use.', False, 'MIST')])
src(s, 'Hytera proposal; 450 MHz Alliance business model material')

# ================================================================ 21 路线图
s = newslide()
page(s, 21, 'Next steps', 'Roadmap: pilot first, then scale',
     'One typhoon-belt province, two anchor customers, and the data Globe needs for a national decision.')
STEPS = [('01', 'Confirm the spectrum', 'Band, bandwidth and licence term', '\u7ebf\u7a3f_\u624b\u53f0_2.png'),
         ('02', 'Pick the pilot region', 'One province, plus two anchors', '\u83f2\u5f8b\u5bbe\u89e3\u51b3\u65b9\u6848\u5305\u88c5_slide11_1.png'),
         ('03', 'Build the pilot', 'Sites, platform, terminals, meters', '\u901a\u4fe1\u94c1\u5854_3.png'),
         ('04', 'Prove the case', 'Coverage, availability and cost data', '\u529e\u516c\u697c.png'),
         ('05', 'Scale nationwide', 'Emergency, transport, smart city', '\u57ce\u5e02\u697c\u7fa4.png')]
for i, (k, t, d, ic) in enumerate(STEPS):
    x = X0 + i * 168
    box(s, x, 122, 186, 96, C['NAVY'] if i == 0 else C['NAVY_L'], shape=MSO_SHAPE.PENTAGON)
    tx(s, x + 14, 132, 150, 16, k, 12, True, 'LIME')
    tx(s, x + 14, 150, 150, 30, t, 10.5, True, 'WHITE', line=1.15)
    tx(s, x + 14, 182, 152, 30, d, 8.5, False, 'MIST', line=1.25)
    box(s, x, 226, 178, 44, C['PAPER'])
    fit(s, illu(ic), x + 68, 232, 42, 32)
box(s, X0, 284, CW, 96, C['CARD'])
box(s, X0, 284, 3.5, 96, C['GREEN'])
tx(s, X0 + 20, 294, 300, 20, 'What we need from Globe', 12.5, True, 'NAVY')
NEEDS = [('Spectrum parameters', 'The band and bandwidth, so the coverage and capacity modelling is exact.'),
         ('A candidate province', 'One region, ideally in the typhoon belt, where the gap is visible.'),
         ('Anchor introductions', 'One electric cooperative and one water district willing to trial.')]
for i, (k, d) in enumerate(NEEDS):
    x = X0 + 20 + i * 280
    box(s, x, 322, 6, 6, C['LIME'])
    tx(s, x + 16, 318, 248, 16, k, 10.5, True, 'NAVY')
    tx(s, x + 16, 338, 250, 36, d, 9.5, False, 'STEEL', line=1.3)
box(s, X0, 392, CW, 36, C['GREEN_L'])
rich(s, X0 + 18, 392, CW - 36, 36,
     [('We are not proposing a build.  ', True, 'GREEN'),
      ('The next step is a joint pilot design, so Globe can test the assumptions in its own terrain.',
       False, 'STEEL')], 11, anchor=MID)
concl(s, [('The technical case is settled.  ', True, 'WHITE'),
          ('What decides the outcome is whether demand is contracted before build-out.', False, 'MIST')])
src(s, 'Hytera proposal')

# ================================================================ 22 尾页
s = newslide()
box(s, 0, 0, W, H, C['WHITE'])
box(s, 0, 0, W, 6, C['NAVY'])
box(s, 0, 0, 320, 6, C['LIME'])
tx(s, X0, 140, 380, 56, 'Thank you', 40, True, 'NAVY')
box(s, X0, 216, 66, 4, C['LIME'])
tx(s, X0, 238, 380, 22, 'Mission First  \u00b7  Safety Always', 14, True, 'GREEN')
tx(s, X0, 272, 400, 44, 'Hytera \u00d7 Globe Telecom  \u00b7  450 MHz for critical infrastructure and remote communities',
   11, False, 'STEEL', line=1.35)
tx(s, X0, 332, 400, 60,
   'Next step: confirm the spectrum parameters and nominate a pilot region, and we will design the pilot with you.',
   11.5, False, 'STEEL', line=1.4)
box(s, 486, 96, 426, 300, C['PAPER'])
fit(s, illu('Hytera 450M Product Introduction PPT_202609_slide9_5.png'), 512, 116, 374, 200)
tx(s, 512, 330, 374, 46, 'Terminal, network and application - one architecture, six industries.',
   10, True, 'NAVY', CEN, line=1.3)
box(s, 486, 96, 426, 300, None, C['LINE'], 1.0)
box(s, 0, 432, W, 108, C['CARD'])
box(s, 0, 428, W, 4, C['LIME'])
for i, n in enumerate(['\u7535\u529b\u62a2\u4fee.png', '\u6c34\u5e93\u5927\u575d.png',
                       '\u8b66\u5bdf\u5c40.png', '\u6d88\u9632\u7ad9\u706b\u707e.png',
                       '\u4ea4\u901a\u67a2\u7ebd.png', '\u8239\u8238.png' if False else '\u8239\u8236.png',
                       '\u96c6\u88c5\u7bb1\u6307\u6325\u6240.png']):
    fit(s, illu(n), 62 + i * 122, 442, 104, 88)
tx(s, X0, 528, 700, 12, 'Hytera Communications Corporation Limited', 8, False, 'SILVER')
