Option Strict Off
Option Explicit On
Module ATSApiVB
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' ATS-API Visual Basic Interface
	' Version 5.7.22
	' Copyright © 2003-2010 Alazar Technologies Inc.
	' All Rights Reserved
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' This module contains function declarations and constants exported
	' by the AlazarTech API (ATSApiVB.dll) to control AlazarTech digitizers.
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' Constants
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Return codes
	
	Public Const ApiSuccess As Short = 512
	Public Const ApiFailed As Short = 513
	Public Const ApiAccessDenied As Short = 514
	Public Const ApiDmaChannelUnavailable As Short = 515
	Public Const ApiDmaChannelInvalid As Short = 516
	Public Const ApiDmaChannelTypeError As Short = 517
	Public Const ApiDmaInProgress As Short = 518
	Public Const ApiDmaDone As Short = 519
	Public Const ApiDmaPaused As Short = 520
	Public Const ApiDmaNotPaused As Short = 521
	Public Const ApiDmaCommandInvalid As Short = 522
	Public Const ApiDmaManReady As Short = 523
	Public Const ApiDmaManNotReady As Short = 524
	Public Const ApiDmaInvalidChannelPriority As Short = 525
	Public Const ApiDmaManCorrupted As Short = 526
	Public Const ApiDmaInvalidElementIndex As Short = 527
	Public Const ApiDmaNoMoreElements As Short = 528
	Public Const ApiDmaSglInvalid As Short = 529
	Public Const ApiDmaSglQueueFull As Short = 530
	Public Const ApiNullParam As Short = 531
	Public Const ApiInvalidBusIndex As Short = 532
	Public Const ApiUnsupportedFunction As Short = 533
	Public Const ApiInvalidPciSpace As Short = 534
	Public Const ApiInvalidIopSpace As Short = 535
	Public Const ApiInvalidSize As Short = 536
	Public Const ApiInvalidAddress As Short = 537
	Public Const ApiInvalidAccessType As Short = 538
	Public Const ApiInvalidIndex As Short = 539
	Public Const ApiMuNotReady As Short = 540
	Public Const ApiMuFifoEmpty As Short = 541
	Public Const ApiMuFifoFull As Short = 542
	Public Const ApiInvalidRegister As Short = 543
	Public Const ApiDoorbellClearFailed As Short = 544
	Public Const ApiInvalidUserPin As Short = 545
	Public Const ApiInvalidUserState As Short = 546
	Public Const ApiEepromNotPresent As Short = 547
	Public Const ApiEepromTypeNotSupported As Short = 548
	Public Const ApiEepromBlank As Short = 549
	Public Const ApiConfigAccessFailed As Short = 550
	Public Const ApiInvalidDeviceInfo As Short = 551
	Public Const ApiNoActiveDriver As Short = 552
	Public Const ApiInsufficientResources As Short = 553
	Public Const ApiObjectAlreadyAllocated As Short = 554
	Public Const ApiAlreadyInitialized As Short = 555
	Public Const ApiNotInitialized As Short = 556
	Public Const ApiBadConfigRegEndianMode As Short = 557
	Public Const ApiInvalidPowerState As Short = 558
	Public Const ApiPowerDown As Short = 559
	Public Const ApiFlybyNotSupported As Short = 560
	Public Const ApiNotSupportThisChannel As Short = 561
	Public Const ApiNoAction As Short = 562
	Public Const ApiHSNotSupported As Short = 563
	Public Const ApiVPDNotSupported As Short = 564
	Public Const ApiVpdNotEnabled As Short = 565
	Public Const ApiNoMoreCap As Short = 566
	Public Const ApiInvalidOffset As Short = 567
	Public Const ApiBadPinDirection As Short = 568
	Public Const ApiPciTimeout As Short = 569
	Public Const ApiDmaChannelClosed As Short = 570
	Public Const ApiDmaChannelError As Short = 571
	Public Const ApiInvalidHandle As Short = 572
	Public Const ApiBufferNotReady As Short = 573
	Public Const ApiInvalidData As Short = 574
	Public Const ApiDoNothing As Short = 575
	Public Const ApiDmaSglBuildFailed As Short = 576
	Public Const ApiPMNotSupported As Short = 577
	Public Const ApiInvalidDriverVersion As Short = 578
	Public Const ApiWaitTimeout As Short = 579
	Public Const ApiWaitCanceled As Short = 580
	Public Const ApiBufferTooSmall As Short = 581
	Public Const ApiBufferOverflow As Short = 582
	Public Const ApiInvalidBuffer As Short = 583
	Public Const ApiInvalidRecordsPerBuffer As Short = 584
	Public Const ApiDmaPending As Short = 585
	Public Const ApiLockAndProbePagesFailed As Short = 586
	Public Const ApiWaitAbandoned As Short = 587
	Public Const ApiWaitFailed As Short = 588
	Public Const ApiTransferComplete As Short = 589
	Public Const ApiPllNotLocked As Short = 590
	Public Const ApiNotSupportedInDualChannelMode As Short = 591
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board types (AlazarGetBoardKind)
	
	Public Const ATS_NONE As Short = 0
	Public Const ATS850 As Short = 1
	Public Const ATS310 As Short = 2
	Public Const ATS330 As Short = 3
	Public Const ATS855 As Short = 4
	Public Const ATS315 As Short = 5
	Public Const ATS335 As Short = 6
	Public Const ATS460 As Short = 7
	Public Const ATS860 As Short = 8
	Public Const ATS660 As Short = 9
	Public Const ATS665 As Short = 10
	Public Const ATS9462 As Short = 11
	Public Const ATS9870 As Short = 13
	Public Const ATS9350 As Short = 14
	Public Const ATS9325 As Short = 15
	Public Const ATS9440 As Short = 16
	Public Const ATS9410 As Short = 17
	Public Const ATS9351 As Short = 18
	Public Const ATS9310 As Short = 19
	Public Const ATS9461 As Short = 20
	Public Const ATS9850 As Short = 21
	Public Const ATS_LAST As Short = 22
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Clock sources (AlazarSetCaptureClock)
	
	Public Const INTERNAL_CLOCK As Short = 1
	Public Const EXTERNAL_CLOCK As Short = 2
	Public Const FAST_EXTERNAL_CLOCK As Short = 2
	Public Const MEDIUM_EXTERNAL_CLOCK As Short = 3
	Public Const SLOW_EXTERNAL_CLOCK As Short = 4
	Public Const EXTERNAL_CLOCK_AC As Short = 5
	Public Const EXTERNAL_CLOCK_DC As Short = 6
	Public Const EXTERNAL_CLOCK_10MHz_REF As Short = 7
	Public Const INTERNAL_CLOCK_DIV_5 As Short = 16
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Sample rates (AlazarSetCaptureClock)
	
	Public Const SAMPLE_RATE_1KSPS As Integer = &H1
	Public Const SAMPLE_RATE_2KSPS As Integer = &H2
	Public Const SAMPLE_RATE_5KSPS As Integer = &H4
	Public Const SAMPLE_RATE_10KSPS As Integer = &H8
	Public Const SAMPLE_RATE_20KSPS As Integer = &HA
	Public Const SAMPLE_RATE_50KSPS As Integer = &HC
	Public Const SAMPLE_RATE_100KSPS As Integer = &HE
	Public Const SAMPLE_RATE_200KSPS As Integer = &H10
	Public Const SAMPLE_RATE_500KSPS As Integer = &H12
	Public Const SAMPLE_RATE_1MSPS As Integer = &H14
	Public Const SAMPLE_RATE_2MSPS As Integer = &H18
	Public Const SAMPLE_RATE_5MSPS As Integer = &H1A
	Public Const SAMPLE_RATE_10MSPS As Integer = &H1C
	Public Const SAMPLE_RATE_20MSPS As Integer = &H1E
	Public Const SAMPLE_RATE_25MSPS As Integer = &H21
	Public Const SAMPLE_RATE_50MSPS As Integer = &H22
	Public Const SAMPLE_RATE_100MSPS As Integer = &H24
	Public Const SAMPLE_RATE_125MSPS As Integer = &H25
	Public Const SAMPLE_RATE_160MSPS As Integer = &H26
	Public Const SAMPLE_RATE_180MSPS As Integer = &H27
	Public Const SAMPLE_RATE_200MSPS As Integer = &H28
	Public Const SAMPLE_RATE_250MSPS As Integer = &H2B
	Public Const SAMPLE_RATE_500MSPS As Integer = &H30
	Public Const SAMPLE_RATE_1GSPS As Integer = &H35
	Public Const SAMPLE_RATE_2GSPS As Integer = &H3A
	Public Const SAMPLE_RATE_USER_DEF As Integer = &H40
	Public Const PLL_10MHZ_REF_100MSPS_BASE As Integer = &H5F5E100
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Clock edges (AlazarSetCaptureClock)
	
	Public Const CLOCK_EDGE_RISING As Short = 0
	Public Const CLOCK_EDGE_FALLING As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Dummy clock (AlazarSetClockSwitchOver)
	
	Public Const CSO_DISABLE As Integer = 0
	Public Const CSO_ENABLE_DUMMY_CLOCK As Integer = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Data skipping (AlazarConfigureSampleSkipping)
	
	Public Const SSM_DISABLE As Integer = 0
	Public Const SSM_ENABLE As Integer = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Channels (AlazarInputControl, AlazarBeforeAsyncRead, AlazarStartAutoDMA)
	
	Public Const CHANNEL_ALL As Short = 0
	Public Const CHANNEL_A As Short = 1
	Public Const CHANNEL_B As Short = 2
	Public Const CHANNEL_C As Short = 4
	Public Const CHANNEL_D As Short = 8
	Public Const CHANNEL_E As Short = 16
	Public Const CHANNEL_F As Short = 32
	Public Const CHANNEL_G As Short = 64
	Public Const CHANNEL_H As Short = 128
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Impedances (AlazarInputControl)
	
	Public Const IMPEDANCE_1M_OHM As Short = 1
	Public Const IMPEDANCE_50_OHM As Short = 2
	Public Const IMPEDANCE_75_OHM As Short = 4
	Public Const IMPEDANCE_300_OHM As Short = 8
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Couplings (AlazarInputControl)
	
	Public Const AC_COUPLING As Short = 1
	Public Const DC_COUPLING As Short = 2
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Input ranges (AlazarInputControl)
	
	Public Const INPUT_RANGE_PM_20_MV As Short = &H1
	Public Const INPUT_RANGE_PM_40_MV As Short = &H2
	Public Const INPUT_RANGE_PM_50_MV As Short = &H3
	Public Const INPUT_RANGE_PM_80_MV As Short = &H4
	Public Const INPUT_RANGE_PM_100_MV As Short = &H5
	Public Const INPUT_RANGE_PM_200_MV As Short = &H6
	Public Const INPUT_RANGE_PM_400_MV As Short = &H7
	Public Const INPUT_RANGE_PM_500_MV As Short = &H8
	Public Const INPUT_RANGE_PM_800_MV As Short = &H9
	Public Const INPUT_RANGE_PM_1_V As Short = &HA
	Public Const INPUT_RANGE_PM_2_V As Short = &HB
	Public Const INPUT_RANGE_PM_4_V As Short = &HC
	Public Const INPUT_RANGE_PM_5_V As Short = &HD
	Public Const INPUT_RANGE_PM_8_V As Short = &HE
	Public Const INPUT_RANGE_PM_10_V As Short = &HF
	Public Const INPUT_RANGE_PM_20_V As Short = &H10
	Public Const INPUT_RANGE_PM_40_V As Short = &H11
	Public Const INPUT_RANGE_PM_16_V As Short = &H12
	Public Const INPUT_RANGE_HIFI As Short = &H20
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Trigger engines (AlazarSetTriggerOperation)
	
	Public Const TRIG_ENGINE_J As Short = 0
	Public Const TRIG_ENGINE_K As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Trigger operations (AlazarSetTriggerOperation)
	
	Public Const TRIG_ENGINE_OP_J As Short = 0
	Public Const TRIG_ENGINE_OP_K As Short = 1
	Public Const TRIG_ENGINE_OP_J_OR_K As Short = 2
	Public Const TRIG_ENGINE_OP_J_AND_K As Short = 3
	Public Const TRIG_ENGINE_OP_J_XOR_K As Short = 4
	Public Const TRIG_ENGINE_OP_J_AND_NOT_K As Short = 5
	Public Const TRIG_ENGINE_OP_NOT_J_AND_K As Short = 6
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Trigger sources (AlazarSetTriggerOperation)
	
	Public Const TRIG_CHAN_A As Short = 0
	Public Const TRIG_CHAN_B As Short = 1
	Public Const TRIG_EXTERNAL As Short = 2
	Public Const TRIG_DISABLE As Short = 3
	Public Const TRIG_CHAN_C As Short = 4
	Public Const TRIG_CHAN_D As Short = 5
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Trigger slopes (AlazarSetTriggerOperation)
	
	Public Const TRIGGER_SLOPE_POSITIVE As Short = 1
	Public Const TRIGGER_SLOPE_NEGATIVE As Short = 2
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' External trigger range (AlazarSetExternalTrigger)
	
	Public Const ETR_5V As Short = 0
	Public Const ETR_1V As Short = 1
	Public Const ETR_TTL As Short = 2
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' LED control (AlazarSetLED)
	
	Public Const LED_OFF As Short = 0
	Public Const LED_ON As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Bandwidth control (AlazarSetBWLimit)
	
	Public Const BW_LIMIT_OFF As Short = 0
	Public Const BW_LIMIT_ON As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Power control (AlazarSleepDevice)
	
	Public Const POWER_OFF As Short = 0
	Public Const POWER_ON As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Timestamp control (AlazarResetTimeStamp)
	
	Public Const TIMESTAMP_RESET_FIRSTTIME_ONLY As Short = 0
	Public Const TIMESTAMP_RESET_ALWAYS As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Sync AutoDMA events (AlazarEvents)
	
	Public Const SW_EVENTS_OFF As Short = 0
	Public Const SW_EVENTS_ON As Short = 1
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' AutoDMA flags (AlazarBeforeAsyncRead, AlazarStartAutoDMA)
	
	Public Const ADMA_EXTERNAL_STARTCAPTURE As Integer = &H1
	Public Const ADMA_TRADITIONAL_MODE As Integer = &H0
	Public Const ADMA_ENABLE_RECORD_HEADERS As Integer = &H8
	Public Const ADMA_SINGLE_DMA_CHANNEL As Integer = &H10
	Public Const ADMA_ALLOC_BUFFERS As Integer = &H20
	Public Const ADMA_CONTINUOUS_MODE As Integer = &H100
	Public Const ADMA_NPT As Integer = &H200
	Public Const ADMA_TRIGGERED_STREAMING As Integer = &H400
	Public Const ADMA_FIFO_ONLY_STREAMING As Integer = &H800
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Sync AutoDMA return codes (AlazarStartAutoDMA ...)
	
	Public Const ADMA_Completed As Integer = 0
	Public Const ADMA_Buffer1Invalid As Integer = 1
	Public Const ADMA_Buffer2Invalid As Integer = 2
	Public Const ADMA_BoardHandleInvalid As Integer = 3
	Public Const ADMA_InternalBuffer1Invalid As Integer = 4
	Public Const ADMA_InternalBuffer2Invalid As Integer = 5
	Public Const ADMA_OverFlow As Integer = 6
	Public Const ADMA_InvalidChannel As Integer = 7
	Public Const ADMA_DMAInProgress As Integer = 8
	Public Const ADMA_UseHeaderNotSet As Integer = 9
	Public Const ADMA_HeaderNotValid As Integer = 10
	Public Const ADMA_InvalidRecsPerBuffer As Integer = 11
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' AutoDMA header items (AlazarGetAutoDMAHeaderValue)
	
	Public Const ADMA_CLOCKSOURCE As Integer = &H1
	Public Const ADMA_CLOCKEDGE As Integer = &H2
	Public Const ADMA_SAMPLERATE As Integer = &H3
	Public Const ADMA_INPUTRANGE As Integer = &H4
	Public Const ADMA_INPUTCOUPLING As Integer = &H5
	Public Const ADMA_IMPUTIMPEDENCE As Integer = &H6
	Public Const ADMA_EXTTRIGGERED As Integer = &H7
	Public Const ADMA_CHA_TRIGGERED As Integer = &H8
	Public Const ADMA_CHB_TRIGGERED As Integer = &H9
	Public Const ADMA_TIMEOUT As Integer = &HA
	Public Const ADMA_THISCHANTRIGGERED As Integer = &HB
	Public Const ADMA_SERIALNUMBER As Integer = &HC
	Public Const ADMA_SYSTEMNUMBER As Integer = &HD
	Public Const ADMA_BOARDNUMBER As Integer = &HE
	Public Const ADMA_WHICHCHANNEL As Integer = &HF
	Public Const ADMA_SAMPLERESOLUTION As Integer = &H10
	Public Const ADMA_DATAFORMAT As Integer = &H11
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' FGPA download (AlazarParseFPGAName)
	
	Public Const FPGA_GETFIRST As Integer = &HFFFFFFFF
	Public Const FPGA_GETNEXT As Integer = &HFFFFFFFE
	Public Const FPGA_GETLAST As Integer = &HFFFFFFFC
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Auxilary I/O (AlazarConfigureAuxIO)
	
	Public Const AUX_OUT_TRIGGER As Integer = 0
	Public Const AUX_OUT_PACER As Integer = 2
	Public Const AUX_OUT_BUSY As Integer = 4
	Public Const AUX_OUT_CLOCK As Integer = 6
	Public Const AUX_OUT_RESERVED As Integer = 8
	Public Const AUX_OUT_CAPTURE_ALMOST_DONE As Integer = 10
	Public Const AUX_OUT_AUXILIARY As Integer = 12
	Public Const AUX_OUT_SERIAL_DATA As Integer = 14
	Public Const AUX_OUT_TRIGGER_ENABLE As Integer = 16
	Public Const AUX_IN_TRIGGER_ENABLE As Integer = 1
	Public Const AUX_IN_DIGITAL_TRIGGER As Integer = 3
	Public Const AUX_IN_GATE As Integer = 5
	Public Const AUX_IN_CAPTURE_ON_DEMAND As Integer = 7
	Public Const AUX_IN_RESET_TIMESTAMP As Integer = 9
	Public Const AUX_IN_SLOW_EXTERNAL_CLOCK As Integer = 11
	Public Const AUX_INPUT_AUXILIARY As Integer = 13
	Public Const AUX_INPUT_SERIAL_DATA As Integer = 15
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board capabilities (AlazarQueryCapability)
	
	Public Const NUMBER_OF_RECORDS As Integer = &H10000001
	Public Const PRETRIGGER_AMOUNT As Integer = &H10000002
	Public Const RECORD_LENGTH As Integer = &H10000003
	Public Const TRIGGER_ENGINE As Integer = &H10000004
	Public Const TRIGGER_DELAY As Integer = &H10000005
	Public Const TRIGGER_TIMEOUT As Integer = &H10000006
	Public Const SAMPLE_RATE As Integer = &H10000007
	Public Const CONFIGURATION_MODE As Integer = &H10000008
	Public Const DATA_WIDTH As Integer = &H10000009
	Public Const SAMPLE_SIZE As Integer = DATA_WIDTH
	Public Const AUTO_CALIBRATE As Integer = &H1000000A
	Public Const TRIGGER_XXXXX As Integer = &H1000000B
	Public Const CLOCK_SOURCE As Integer = &H1000000C
	Public Const CLOCK_SLOPE As Integer = &H1000000D
	Public Const IMPEDANCE As Integer = &H1000000E
	Public Const INPUT_RANGE As Integer = &H1000000F
	Public Const COUPLING As Integer = &H10000010
	Public Const MAX_TIMEOUTS_ALLOWED As Integer = &H10000011
	Public Const OPERATING_MODE As Integer = &H10000012
	Public Const CLOCK_DECIMATION_EXTERNAL As Integer = &H10000013
	Public Const LED_CONTROL As Integer = &H10000014
	Public Const ATTENUATOR_RELAY As Integer = &H10000018
	Public Const EXT_TRIGGER_COUPLING As Integer = &H1000001A
	Public Const EXT_TRIGGER_ATTENUATOR_RELAY As Integer = &H1000001C
	Public Const TRIGGER_ENGINE_SOURCE As Integer = &H1000001E
	Public Const TRIGGER_ENGINE_SLOPE As Integer = &H10000020
	Public Const GET_SERIAL_NUMBER As Integer = &H10000024
	Public Const GET_FIRST_CAL_DATE As Integer = &H10000025
	Public Const GET_LATEST_CAL_DATE As Integer = &H10000026
	Public Const GET_LATEST_TEST_DATE As Integer = &H10000027
	Public Const GET_LATEST_CAL_DATE_MONTH As Integer = &H1000002D
	Public Const GET_LATEST_CAL_DATE_DAY As Integer = &H1000002E
	Public Const GET_LATEST_CAL_DATE_YEAR As Integer = &H1000002F
	Public Const GET_PCIE_LINK_SPEED As Integer = &H10000030
	Public Const GET_PCIE_LINK_WIDTH As Integer = &H10000031
	Public Const MEMORY_SIZE As Integer = &H1000002A
	Public Const BOARD_TYPE As Integer = &H1000002B
	Public Const ASOPC_TYPE As Integer = &H1000002C
	Public Const GET_BOARD_OPTIONS_LOW As Integer = &H10000037
	Public Const GET_BOARD_OPTIONS_HIGH As Integer = &H10000038
	Public Const OPTION_STREAMING_DMA As Integer = 1
	Public Const OPTION_EXTERNAL_CLOCK As Integer = 2
	Public Const OPTION_DUAL_PORT_MEMORY As Integer = 4
	Public Const OPTION_180MHZ_OSCILLATOR As Integer = 8
	Public Const OPTION_LVTTL_EXT_CLOCK As Integer = 16
	Public Const OPTION_SW_SPI As Integer = 32
	Public Const TRANSFER_RECORD_OFFSET As Integer = &H10000032
	Public Const TRANSFER_NUM_OF_RECORDS As Integer = &H10000033
	Public Const TRANSFER_MAPPING_RATIO As Integer = &H10000034
	Public Const TRIGGER_ADDRESS_AND_TIMESTAMP As Integer = &H10000035
	Public Const MASTER_SLAVE_INDEPENDENT As Integer = &H10000036
	Public Const TRIGGERED As Integer = &H10000040
	Public Const BUSY As Integer = &H10000041
	Public Const WHO_TRIGGERED As Integer = &H10000042
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board parameters (AlazarSetParameter, AlazarGetParameter)
	
	Public Const SETGET_ASYNC_BUFFSIZE_BYTES As Integer = &H10000039
	Public Const SETGET_ASYNC_BUFFCOUNT As Integer = &H10000040
	Public Const SET_DATA_FORMAT As Integer = &H10000041
	Public Const GET_DATA_FORMAT As Integer = &H10000042
	Public Const DATA_FORMAT_UNSIGNED As Integer = 0
	Public Const DATA_FORMAT_SIGNED As Integer = 1
	Public Const SET_SINGLE_CHANNEL_MODE As Integer = &H10000043
	Public Const GET_SAMPLES_PER_TIMESTAMP_CLOCK As Integer = &H10000044
	Public Const GET_RECORDS_CAPTURED As Integer = &H10000045
	Public Const GET_MAX_PRETRIGGER_SAMPLES As Integer = &H10000046
	Public Const SET_ADC_MODE As Integer = &H10000047
	Public Const ECC_MODE As Integer = &H10000048
	Public Const ECC_DISABLE As Integer = 0
	Public Const ECC_ENABLE As Integer = 1
	Public Const GET_AUX_INPUT_LEVEL As Integer = &H10000049
	Public Const AUX_INPUT_LOW As Integer = 0
	Public Const AUX_INPUT_HIGH As Integer = 1
	Public Const EXT_TRIGGER_IMPEDANCE As Integer = &H10000065
	Public Const EXT_TRIG_50_OHMS As Integer = 0
	Public Const EXT_TRIG_300_OHMS As Integer = 1
	Public Const GET_ASYNC_BUFFERS_PENDING As Integer = &H10000050
	Public Const GET_ASYNC_BUFFERS_PENDING_FULL As Integer = &H10000051
	Public Const GET_ASYNC_BUFFERS_PENDING_EMPTY As Integer = &H10000052
	Public Const ACF_SAMPLES_PER_RECORD As Integer = &H10000060
	Public Const ACF_RECORDS_TO_AVERAGE As Integer = &H10000061
	Public Const ACF_MODE As Integer = &H10000062
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board parameters (AlazarSetParameterUL, AlazarGetParameterUL)
	
	Public Const SEND_DAC_VALUE As Integer = &H10000021
	Public Const SLEEP_DEVICE As Integer = &H10000022
	Public Const GET_DAC_VALUE As Integer = &H10000023
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' Function declarations
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board identification
	
	Public Declare Function AlazarGetSystemHandle Lib "ATSApiVB.dll" (ByVal sid As Short) As Integer
	Public Declare Function AlazarNumOfSystems Lib "ATSApiVB.dll" () As Integer
	Public Declare Function AlazarBoardsFound Lib "ATSApiVB.dll" () As Integer
	Public Declare Function AlazarGetBoardKind Lib "ATSApiVB.dll" (ByVal boardHandle As Integer) As Integer
	Public Declare Function AlazarBoardsInSystemBySystemID Lib "ATSApiVB.dll" (ByVal sid As Integer) As Integer
	Public Declare Function AlazarBoardsInSystemByHandle Lib "ATSApiVB.dll" (ByVal systemHandle As Integer) As Integer
	Public Declare Function AlazarGetBoardBySystemID Lib "ATSApiVB.dll" (ByVal sid As Integer, ByVal brdNum As Integer) As Integer
	Public Declare Function AlazarGetBoardBySystemHandle Lib "ATSApiVB.dll" (ByVal systemHandle As Integer, ByVal brdNum As Integer) As Integer
	Public Declare Function AlazarOpen Lib "ATSApiVB.dll" (ByVal Name As String) As Integer
	Public Declare Sub AlazarClose Lib "ATSApiVB.dll" (ByVal h As Integer)
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Status and information
	
	Public Declare Function AlazarGetCPLDVersion Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef Major As Byte, ByRef Minor As Byte) As Integer
	Public Declare Function AlazarGetSDKVersion Lib "ATSApiVB.dll" (ByRef Major As Byte, ByRef Minor As Byte, ByRef Revision As Byte) As Integer
	Public Declare Function AlazarGetDriverVersion Lib "ATSApiVB.dll" (ByRef Major As Byte, ByRef Minor As Byte, ByRef Revision As Byte) As Integer
	Public Declare Function AlazarBusy Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarTriggered Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarGetStatus Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarGetChannelInfo Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef MemSize As Integer, ByRef SampleSize As Byte) As Integer
	Public Declare Function AlazarTriggerTimedOut Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarGetTriggerTimestamp Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal record As Integer, ByRef Timestamp_samples As Decimal) As Object
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Board configuration
	
	'UPGRADE_NOTE: Rate was upgraded to Rate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Declare Function AlazarSetCaptureClock Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal Source As Integer, ByVal Rate_Renamed As Integer, ByVal Edge As Integer, ByVal Decimation As Integer) As Integer
	Public Declare Function AlazarInputControl Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Byte, ByVal COUPLING As Integer, ByVal InputRange As Integer, ByVal IMPEDANCE As Integer) As Integer
	Public Declare Function AlazarSetTriggerOperation Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal TriggerOperation As Integer, ByVal TriggerEngine1 As Integer, ByVal Source1 As Integer, ByVal Slope1 As Integer, ByVal Level1 As Integer, ByVal TriggerEngine2 As Integer, ByVal Source2 As Integer, ByVal Slope2 As Integer, ByVal Level2 As Integer) As Integer
	Public Declare Function AlazarSetExternalTrigger Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal COUPLING As Integer, ByVal Range As Integer) As Integer
	Public Declare Function AlazarSetTriggerDelay Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal Delay As Integer) As Integer
	Public Declare Function AlazarSetTriggerTimeOut Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal to_ns As Integer) As Integer
	Public Declare Function AlazarSetLED Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal state As Short) As Integer
	Public Declare Function AlazarSetBWLimit Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Integer, ByVal enable As Integer) As Integer
	Public Declare Function AlazarSleepDevice Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal state As Integer) As Integer
	Public Declare Function AlazarReadRegister Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal a As Integer, ByRef B As Integer, ByVal pswrd As Integer) As Integer
	Public Declare Function AlazarWriteRegister Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal a As Integer, ByVal B As Integer, ByVal pswrd As Integer) As Integer
	Public Declare Function AlazarResetTimeStamp Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal resetFlag As Integer) As Integer
	Public Declare Function AlazarMemoryTest Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef errors As Integer) As Integer
	Public Declare Function AlazarAutoCalibrate Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarQueryCapability Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal request As Integer, ByVal valueIn As Integer, ByRef value As Integer) As Integer
	Public Declare Function AlazarConfigureAuxIO Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal mode As Integer, ByVal options As Integer) As Integer
	Public Declare Function AlazarSetParameter Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Byte, ByVal parameter As Integer, ByVal value As Integer) As Integer
	Public Declare Function AlazarGetParameter Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Byte, ByVal parameter As Integer, ByRef value As Integer) As Integer
	Public Declare Function AlazarSetParameterUL Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Byte, ByVal parameter As Integer, ByVal value As Integer) As Integer
	Public Declare Function AlazarGetParameterUL Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Byte, ByVal parameter As Integer, ByRef value As Integer) As Integer
	Public Declare Function AlazarSetExternalClockLevel Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal level_percent As Single) As Object
	Public Declare Function AlazarSetClockSwitchOver Lib "ATSApiVB.dll" (ByVal h As Object, ByVal mode As Object, ByVal dummClockOnTime_ns As Object, ByVal reserved As Object) As Object
	Public Declare Function AlazarSetTriggerOperationForScanning Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal slope As Integer, ByVal level As Integer, ByVal options As Integer) As Object
	Public Declare Function AlazarConfigureSampleSkipping Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal mode As Integer, ByVal sampleClocksPerTrigger As Integer, ByRef skipTable As Short) As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' General acquisition
	
	Public Declare Function AlazarSetRecordSize Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal PreSize As Integer, ByVal PostSize As Integer) As Integer
	Public Declare Function AlazarForceTrigger Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarForceTriggerEnable Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarStartCapture Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Single-port acquisition
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function AlazarRead Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Integer, ByRef buffer As Short, ByVal ElementSize As Short, ByVal record As Integer, ByVal TransferOffset As Integer, ByVal TransferLength As Integer) As Integer
	Public Declare Function AlazarDetectMultipleRecord Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarGetTriggerAddress Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal record As Integer, ByRef TriggerAddress As Integer, ByRef TimeStampHighPart As Integer, ByRef TimeStampLowPart As Integer) As Integer
	Public Declare Function AlazarSetRecordCount Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal Count As Integer) As Integer
	Public Declare Function AlazarAbortCapture Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	Public Declare Function AlazarGetWhoTriggeredBySystemHandle Lib "ATSApiVB.dll" (ByVal systemHandle As Integer, ByVal brdNum As Integer, ByVal recNum As Integer) As Integer
	Public Declare Function AlazarGetWhoTriggeredBySystemID Lib "ATSApiVB.dll" (ByVal sid As Short, ByVal brdNum As Integer, ByVal recNum As Integer) As Integer
	Public Declare Function AlazarMaxSglTransfer Lib "ATSApiVB.dll" (ByVal bt As Short) As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Dual-port synchronous AutoDMA
	
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarStartAutoDMA Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef Buffer1 As Short, ByVal UseHeader As Integer, ByVal ChannelSelect As Integer, ByVal TransferOffset As Short, ByVal TransferLength As Integer, ByVal RecordsPerBuffer As Integer, ByVal RecordCount As Integer, ByRef Error_Renamed As Integer, ByVal r1 As Integer, ByVal r2 As Integer, ByRef r3 As Integer, ByRef r4 As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarGetNextAutoDMABuffer Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef Buffer1 As Short, ByRef Buffer2 As Short, ByRef WhichOne As Integer, ByRef RecordsTransfered As Integer, ByRef Error_Renamed As Integer, ByVal r1 As Integer, ByVal r2 As Integer, ByRef TriggersOccurred As Integer, ByRef r4 As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarGetNextBuffer Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef Buffer1 As Short, ByRef Buffer2 As Short, ByRef WhichOne As Integer, ByRef RecordsTransfered As Integer, ByRef Error_Renamed As Integer, ByVal r1 As Integer, ByVal r2 As Integer, ByRef TriggersOccurred As Integer, ByRef r4 As Integer) As Integer
	Public Declare Function AlazarCloseAUTODma Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarAbortAutoDMA Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef Buffer1 As Short, ByRef Error_Renamed As Integer, ByVal r1 As Integer, ByVal r2 As Integer, ByRef r3 As Integer, ByRef r4 As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarGetAutoDMAHeaderValue Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Integer, ByRef DataBuffer As Short, ByVal record As Integer, ByVal parameter As Integer, ByRef Error_Renamed As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarGetAutoDMAHeaderTimeStamp Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal channel As Integer, ByRef DataBuffer As Short, ByVal record As Integer, ByRef Error_Renamed As Integer) As Double
	Public Declare Function AlazarWaitForBufferReady Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal tms As Integer) As Integer
	Public Declare Function AlazarEvents Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal enable As Integer) As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Dual-port asynchronous AutoDMA
	
	Public Declare Function AlazarBeforeAsyncRead Lib "ATSApiVB.dll" (ByVal h As Integer, ByVal uChannelSelect As Integer, ByVal lTransferOffset As Integer, ByVal uSamplesPerRecord As Integer, ByVal uRecordsPerBuffer As Integer, ByVal uRecordsPerAcquisition As Integer, ByVal uFlags As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarWaitAsyncBufferComplete Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef buffer As Short, ByVal uTimeout_ms As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarWaitNextAsyncBufferComplete Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef buffer As Short, ByVal uBufferLength_bytes As Integer, ByVal uTimeout_ms As Integer) As Integer
	Public Declare Function AlazarAbortAsyncRead Lib "ATSApiVB.dll" (ByVal h As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function AlazarPostAsyncBuffer Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef buffer As Short, ByVal bytesPerBuffer As Integer) As Integer
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' FPGA download
	
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarOEMDownLoadFPGA Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef fileName As Any, ByRef Error_Renamed As Integer) As Short
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarParseFPGAName Lib "ATSApiVB.dll" (ByRef FullPath As Any, ByRef Name As Any, ByRef bType As Integer, ByRef MemSize As Integer, ByRef MajVer As Integer, ByRef MinVer As Integer, ByRef MajRev As Integer, ByRef MinRev As Integer, ByRef Error_Renamed As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarGetOEMFPGAName Lib "ATSApiVB.dll" (ByVal opcodeID As Integer, ByRef FullPath As Any, ByRef Error_Renamed As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarOEMGetWorkingDirectory Lib "ATSApiVB.dll" (ByRef wDir As Any, ByRef Error_Renamed As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarOEMSetWorkingDirectory Lib "ATSApiVB.dll" (ByRef wDir As Any, ByRef Error_Renamed As Integer) As Integer
	'UPGRADE_NOTE: Error was upgraded to Error_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Public Declare Function AlazarStreamCapture Lib "ATSApiVB.dll" (ByVal h As Integer, ByRef buffer As Any, ByVal BufferSize As Integer, ByVal DeviceOption As Integer, ByVal ChannelSelect As Integer, ByRef Error_Renamed As Integer) As Short
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Helper function to convert a null-terminated 'C' string returned
	' by AlazarErrorToText to a VB string
	
	Public Declare Function AlazarErrorToLPCSTR Lib "ATSApiVB.dll"  Alias "AlazarErrorToText"(ByVal retCode As Integer) As Integer
	Declare Function lstrlen Lib "kernel32"  Alias "lstrlenA"(ByVal lpString As Integer) As Integer
	Declare Function lstrcpy Lib "kernel32"  Alias "lstrcpyA"(ByVal lpString1 As String, ByVal lpString2 As Integer) As Integer
	
	Function AlazarErrorToText(ByVal retCode As Integer) As String
		Dim pstr As Integer
		Dim strlen As Integer
		Dim vbstr As String
        'Dim text As String
		Dim pbuffer As Integer
		
		pstr = AlazarErrorToLPCSTR(retCode)
		strlen = lstrlen(pstr)
		vbstr = Space(strlen)
		pbuffer = lstrcpy(vbstr, pstr)
		
		AlazarErrorToText = vbstr & " (" & Str(retCode) & ")"
		
	End Function
End Module