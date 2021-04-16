﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("SWMC_MP.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Grapheme	Prior propbability	Phoneme	FrequencyData	Conditional probability	Conditional Predictability	Example
        '''
        '''Summed frequency data: 8816314
        '''a	0,100086498734051	882394
        '''		ɑː	142530	0,161526483634295	0,197693643510633	straka [str²ɑːka]
        '''		a	720964	0,817054513063325	1	salling [s²alːɪŋ]
        '''		ɑ	17721	0,0200828654773265	0,024579590659173	varunder [vɑrˈɵnːde̞r]
        '''		ɛː	327	0,000370582755549108	0,000453559401024184	alabamas [alabˈɛːmas]
        '''		ɛ̝	149	0,000168858809103416	0,000206667739304598	abstracten [ˈɛːbstrɛ̝kte̞n] [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property g2p_Data() As String
            Get
                Return ResourceManager.GetString("g2p_Data", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Grapheme	FreqData	HighestProb	Grapheme block	FreqData	KIS2K_Conditional propbability	Phoneme	FreqData	K2V_Conditional probability	KIS2V_Conditional probability	KIS2V_Predictability	Examples
        '''
        '''Total KIS count: 34
        '''a	172,195928607792	0,0471400116129678619088168163
        '''			a	57,1970491782514	0,332162610583716
        '''						a	8,11731807427709	0,141918476405661	0,0471400116129678619088168163	1	salling [s²alːɪŋ]
        '''						ɑː	7,80860611590898	0,136521135759537	0,0453472168537417135925838995	0,9619687247015237478230471901	strak [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property GIL2P_Data() As String
            Get
                Return ResourceManager.GetString("GIL2P_Data", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Spelling	Zipf-value
        '''3m	2,440425
        '''7-eleven	2,888415
        '''a	5,750159
        '''à	3,748251
        '''a:et	2,080152
        '''a3	3,074774
        '''a4	3,378165
        '''a4:or	0,9662084
        '''a4-sidor	2,434556
        '''aaby	0,5682684
        '''aachen	2,55504
        '''aagaard	1,589458
        '''aage	1,545992
        '''a-aktier	1,628966
        '''aalborg	2,609661
        '''aalborgs	1,545992
        '''aalto	2,371042
        '''aaltonen	1,785752
        '''aaltos	1,858303
        '''aamulehti	1,7586
        '''aarenstrup	0,7443597
        '''aarhus	1,811306
        '''aarne	1,413366
        '''aarnio	1,497687
        '''aarno	0,5682684
        '''aaron	3,221963
        '''aarons	1,698602
        '''aarre	0,9662084
        '''aaröe	0,5682684
        '''aas	2,04539 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property OLDComparisonCorpus_ArcList() As String
            Get
                Return ResourceManager.GetString("OLDComparisonCorpus_ArcList", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Phon/Graph	PDCS/PhoneticContextConditions	PDS/SpellingContextConditions	ReplPhonemes	Comments	GraphemeCount	ExampleWords_NoPDCS	ExampleWords_PDCS	ExampleWords_PDS	ExampleWords_SilentGraphemes
        '''
        '''NormChars	- : . &apos; / _ *
        '''
        '''
        '''[p]	PDCS				147825
        '''pp					5810	&lt;rapport&gt;, [rapˈɔʈː]				
        '''bb	-p-t				0					
        '''p					141178	&lt;precis&gt;, [pre̞sˈiːs]				
        '''b	-p-t				49	&lt;superbt&gt;, [sʉpˈɛ̝rːpt]				
        '''b		jaco-b-, jako-b-, o-b-s, a-b-s, su-b-s, her-b-st, ry-b-sen, schi-b-sted			788	&lt;absolut&gt;, [apsʊlˈʉːt]				
        '''
        '''[pː]	PDCS				26299 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property p2gRules() As String
            Get
                Return ResourceManager.GetString("p2gRules", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Grapheme	FreqData	HighestProb	Grapheme block	FreqData	KIS2K_Conditional propbability	Phoneme	FreqData	K2V_Conditional probability	KIS2V_Conditional probability	KIS2V_Predictability	Examples
        '''
        '''Total KIS count: 79
        '''∅	12,1102526871335	0,503799295961056
        '''			∅	12,1102526871335	1
        '''						h	6,10113677768832	0,503799295961056	0,503799295961056	1	grehn [ɡrˈeːn]
        '''						e	6,00911590944514	0,496200704038944	0,496200704038944	0,9849174225072768345077707076	modéen [mʊdˈeːn]
        '''
        '''a	95,9750043717041	0,08457741812482037339970 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property PIP2G_Data() As String
            Get
                Return ResourceManager.GetString("PIP2G_Data", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Syllable length	PLD1 transcription	Zipf-value
        '''1	ˈ 1 ɑː	6,051188
        '''1	ˈ 1 a	4,303748
        '''1	ˈ 1 oː s	2,93188
        '''1	ˈ 1 a bː	1,112336
        '''1	ˈ 1 a kː	4,120328
        '''1	ˈ 1 ɑː f	4,153278
        '''1	ˈ 1 ɑː ɡ	3,787984
        '''1	ˈ 1 a ɡː	2,850437
        '''1	ˈ 1 a ŋː n	1,628966
        '''1	ˈ 1 ɑː ɡ s	1,900707
        '''1	ˈ 1 ɑː l	4,113575
        '''1	ˈ 1 a lː m	3,105458
        '''1	ˈ 1 a lː m s	1,880022
        '''1	ˈ 1 ɑː l s	2,748681
        '''1	ˈ 1 ɛ̝ ʝː d s	3,421053
        '''1	ˈ 1 a ʝː n	3,53441
        '''1	ˈ 1 ɛː r	7,194548
        '''1	ˈ 1 ɛː ʂ	2,03809
        '''1	ˈ 1 a ʝː	4,16371
        '''1	ˈ 1 a kː t	3,687194
        '''1	ˈ 1 a lː f	3,388798
        '''1	ˈ 1 a lː [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property PLDComparisonCorpus_ArcList() As String
            Get
                Return ResourceManager.GetString("PLDComparisonCorpus_ArcList", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Position	Phoneme/s	FrequencyData	TransitionalProbability	Z-TransformedTransitionalProbability	Occurences
        '''
        '''0	a b	318,053276356974	0,000563306255710811	-0,178499918340898	608
        '''0	ɑ b	19,9618549860635	3,53545730389551E-05	-0,428188475990092	21
        '''0	a bː	68,9481315304226	0,000122114490551788	-0,387156390334241	133
        '''0	a ɔ	7,99229967076969	1,41552146660076E-05	-0,438214463952554	16
        '''0	a d	629,698412662548	0,00111526301230697	0,0825415575414361	1025
        '''0	ɑ d	8,78861661457469	1,55655768578764E-05	-0,437547449691573	9        ''' [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property PSBP_Matrix_FullLines() As String
            Get
                Return ResourceManager.GetString("PSBP_Matrix_FullLines", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Position	Phoneme/s	FrequencyData	TransitionalProbability	Z-TransformedTransitionalProbability	Occurences
        '''
        '''0	a	20993,162233328	0,0371689794269191	0,677380774521667	33680
        '''0	ɑ	76,5887781083248	0,000135602568407814	-0,751097262764313	108
        '''0	ɑː	7294,73132895869	0,0129155253352218	-0,258140893693447	13016
        '''0	a͡u	547,044993171694	0,000968558422538063	-0,718967890706243	994
        '''0	a͡uː	117,113238874932	0,000207352256795862	-0,748329682199462	162
        '''0	b	41083,7402758161	0,0727399083625631	2,04944821910172	61022
        '''0	ɔ	60 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property PSP_Matrix_FullLines() As String
            Get
                Return ResourceManager.GetString("PSP_Matrix_FullLines", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Stress	Position	Phoneme	TransitionTo	TransitionalPropbability	TransitionalPredictability	JointPosition	JointTransition	PhonemeOccurences	TransitionToOccurences
        '''Ante_Stress	Onset	*	ɪ	0,0331378077921553	1	Ante_Stress-Onset	*-ɪ	79374	3802
        '''Ante_Stress	Onset	*	f	0,0323611633809348	0,976563192831229	Ante_Stress-Onset	*-f	79374	8127
        '''Ante_Stress	Onset	*	s	0,0320626734983373	0,96755566027296	Ante_Stress-Onset	*-s	79374	6198
        '''Ante_Stress	Onset	*	p	0,0320450568371558	0,967024042089525	Ante_Stress-Onset	*-p	79374	73 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property SSPP_Matrix_FullLines() As String
            Get
                Return ResourceManager.GetString("SSPP_Matrix_FullLines", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Word initial onset clusters
        '''Cluster	Frequency
        '''
        '''	120541015
        '''t r	2942453
        '''s	35753909
        '''b	11696866
        '''b ʝ	169494
        '''b l	4143487
        '''b r	2161093
        '''k	12859968
        '''h	29183141
        '''t ɕ	9052
        '''ɕ	3982099
        '''ɧ	3699489
        '''ʂ	84876
        '''k l	1642420
        '''k r	1304248
        '''k n	330588
        '''d	35749809
        '''d ʝ	4697
        '''ʝ	25395460
        '''d n ʝ	24
        '''d r	877557
        '''d v	2821
        '''f	20961151
        '''f ʝ	54330
        '''f l	939339
        '''f n	10691
        '''f r	3799264
        '''ɡ	4413272
        '''ɡ l	656272
        '''ɡ n	45614
        '''ɡ r	1394135
        '''ɡ ʂ	6
        '''v	24860187
        '''k m	172
        '''k ʝ	99
        '''n	13703048
        '''k ʂ	28
        '''k v	938563
        '''l	12519033
        '''l ʝ	60
        '''t	15909533
        ''' [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property SyllabificationClusters() As String
            Get
                Return ResourceManager.GetString("SyllabificationClusters", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
